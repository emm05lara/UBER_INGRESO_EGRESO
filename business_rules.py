# business_rules.py
# ──────────────────────────────────────────────────────────────
# Módulo de reglas de negocio — Operación UBER 2025 / 2026
# ──────────────────────────────────────────────────────────────
# REGLAS DE SIGNO (documentadas aquí para auditoria):
#   RENTA      : valores negativos en Excel → cobro al conductor → abs()
#   FIANZA     : negativo en Excel = ingreso (cobro), positivo = devolución.
#                Se invierte para que positivo = cobrado, negativo = devuelto.
#   MULTA      : negativo = cargo al conductor; positivo = crédito/devolución.
#                NO se usa abs(). Se expone cargo, credito y neto.
#   HOJALATERO : igual que MULTA.
#   DESCUENTOS : igual que MULTA. Es el más crítico: muchos créditos positivos.
# ──────────────────────────────────────────────────────────────
from __future__ import annotations

import warnings
import numpy as np
import pandas as pd

# ═══════════════════════════════════════════════════════════════
# CONFIGURACIÓN DE HOJAS
# ═══════════════════════════════════════════════════════════════
SHEETS = {
    "ingresos": {
        2025: {"name": "UBER 2025", "header_row": 0},
        2026: {"name": "UBER 2026", "header_row": 0},
    },
    "egresos": {
        2025: {"name": "Gastos 2025", "header_row": 0},
        2026: {"name": "Gastos 2026", "header_row": 0},
    },
}

# ═══════════════════════════════════════════════════════════════
# ALIASES DE COLUMNAS — robusto ante variaciones entre años
# ═══════════════════════════════════════════════════════════════
# Clave = nombre canónico interno; valor = lista de aliases en el Excel
# (se usa el primer alias que se encuentre en cada hoja)
COLUMN_ALIASES: dict[str, list[str]] = {
    "semana":             ["SEM"],
    "conductor":          ["CONDUCTOR"],
    "llave":              ["LLAVE", "AUTO"],
    "tag":                ["TAG"],
    "socio":              ["SOCIO"],
    "plataforma":         ["PLATAFORMA"],
    "app":                ["APP"],
    "ganancia":           ["GANANCIA"],
    "renta_raw":          ["RENTA"],
    "sem_pasada":         ["SEM A", "SEM PASADA"],
    "fianza_raw":         ["FIANZA"],
    "multa_raw":          ["MULTA"],
    "hojalatero_raw":     ["HOJALAT", "HOJALATERO"],
    "descuentos_raw":     ["DESC", "DESCUENTOS"],
    "total":              ["TOTAL"],
    "ganancias_totales":  ["GANAN T", "GANANCIAS TOTALES", "GANANCIAS \nTOTALES"],
    "comentarios":        ["COMENTARIOS"],
}

# Columnas numéricas que siempre se convierten
_INGRESOS_NUMERIC = [
    "ganancia", "renta_raw", "sem_pasada", "fianza_raw",
    "multa_raw", "hojalatero_raw", "descuentos_raw",
    "total", "ganancias_totales",
]

_EGRESOS_RENAME = {
    2025: {
        "Semana": "semana",
        "MES": "mes",
        "Fecha": "fecha",
        "Solicitante": "solicitante",
        "CONCEPTO": "concepto",
        "DETALLE": "detalle",
        "CONDUCTOR": "conductor",
        "DETALLE.1": "llave",
        "SOCIO": "socio",
        "METODO DE PAGO": "metodo_pago",
        "REAL": "monto_real",
        "COMERCIO": "comercio",
        "COMENTARIOS": "comentarios",
        "ADICIONAL": "adicional",
    },
    2026: {
        "SEM": "semana",
        "MES": "mes",
        "FECHA": "fecha",
        "CONCEPTO": "concepto",
        "DETALLE": "detalle",
        "CONDUCTOR": "conductor",
        "LLAVE": "llave",
        "SOCIO": "socio",
        "METODO DE PAGO": "metodo_pago",
        "RESPONSABLE": "responsable",
        "REAL": "monto_real",
        "COMERCIO": "comercio",
        "COMENTARIOS": "comentarios",
        "ADICIONAL": "adicional",
        "FOLIO FISCAL": "folio_fiscal",
    },
}

_EGRESOS_NUMERIC = ["monto_real"]


# ═══════════════════════════════════════════════════════════════
# HELPERS
# ═══════════════════════════════════════════════════════════════
def _to_semana(s: pd.Series) -> pd.Series:
    """Convierte semana a float64 (NaN-safe, compatible PyArrow)."""
    return pd.to_numeric(s, errors="coerce")


def _coerce_numeric(s: pd.Series) -> pd.Series:
    """Convierte a numérico: acepta $, comas, paréntesis (negativos)."""
    if s is None or pd.api.types.is_numeric_dtype(s):
        return s
    x = s.astype(str).str.strip()
    x = x.replace({"nan": np.nan, "None": np.nan, "": np.nan, "-": np.nan})
    x = x.str.replace(r"^\((.+)\)$", r"-\1", regex=True)
    x = x.str.replace(r"[^0-9,.\-]", "", regex=True)

    def _fix_commas(val):
        if pd.isna(val):
            return val
        val = str(val)
        if "," not in val:
            return val
        return val.replace(",", "") if "." in val else val.replace(",", ".")

    x = x.apply(_fix_commas)
    return pd.to_numeric(x, errors="coerce")


def _strip_unnamed(df: pd.DataFrame) -> pd.DataFrame:
    return df.loc[:, ~df.columns.astype(str).str.match(r"^Unnamed")]


def _resolve_aliases(raw_cols: list[str], aliases: dict[str, list[str]]) -> dict[str, str]:
    """
    Dado el listado real de columnas del Excel y el diccionario de aliases,
    devuelve {excel_col -> canonical_name} para las columnas encontradas.
    Emite warning si un alias canónico no se encuentra.
    """
    rename_map: dict[str, str] = {}
    for canonical, candidates in aliases.items():
        for candidate in candidates:
            if candidate in raw_cols:
                rename_map[candidate] = canonical
                break
        else:
            # Ningún alias encontrado — warning suave
            if canonical in _INGRESOS_NUMERIC or canonical in ("semana", "conductor"):
                warnings.warn(
                    f"[business_rules] Columna '{canonical}' no encontrada. "
                    f"Aliases buscados: {candidates}",
                    UserWarning,
                    stacklevel=3,
                )
    return rename_map


# ═══════════════════════════════════════════════════════════════
# LÓGICA DE SIGNOS — COLUMNAS CON SIGNO MIXTO
# ═══════════════════════════════════════════════════════════════
def split_cargos_creditos(s: pd.Series) -> tuple[pd.Series, pd.Series, pd.Series]:
    """
    Descompone una Serie con signo mixto en tres métricas:

    Interpretación del Excel:
      - Valor negativo → cargo/cobro al conductor  → aparece como positivo en 'cargo'
      - Valor positivo → crédito/devolución        → aparece tal cual en 'credito'
      - Neto = cargo - credito  (lo que realmente se cobró neto)

    Returns
    -------
    (cargo, credito, neto) — tres Series con el mismo índice que s.
    """
    valores = s.fillna(0)
    cargo   = valores.where(valores < 0, 0).abs()   # neg -> positivo
    credito = valores.where(valores > 0, 0)          # pos -> tal cual
    neto    = cargo - credito
    return cargo, credito, neto


# ═══════════════════════════════════════════════════════════════
# TRANSFORMACIONES — INGRESOS
# ═══════════════════════════════════════════════════════════════
def _invertir_signo_fianza(val):
    """
    Regla fianza: en el Excel '-' es cobro de fianza al conductor (ingreso),
    '+' es devolución. Invertimos para que positivo = cobrado.
    """
    if pd.isna(val):
        return 0.0
    return float(val) * -1


def generar_concepto_ingreso(row: pd.Series) -> str:
    """Genera etiqueta de concepto según qué columnas tienen valor."""
    partes = []
    if abs(row.get("renta_semanal", 0) or 0) > 0:
        partes.append("Renta")
    if (row.get("fianza", 0) or 0) != 0:
        partes.append("Fianza")
    if abs(row.get("multa_neto", 0) or 0) > 0:
        partes.append("Multa")
    if abs(row.get("hojalatero_neto", 0) or 0) > 0:
        partes.append("Hojalatero")
    if abs(row.get("descuentos_neto", 0) or 0) > 0:
        partes.append("Descuento")
    return " + ".join(partes) if partes else "Sin concepto"


def transform_ingresos(df: pd.DataFrame) -> pd.DataFrame:
    """
    Aplica reglas de negocio a un DataFrame de ingresos ya renombrado.

    Columnas generadas por concepto con signo mixto (multa/hojalatero/descuentos):
      *_cargo   : suma de valores negativos convertidos a positivo (cobros reales)
      *_credito : suma de valores positivos (devoluciones/ajustes a favor)
      *_neto    : cargo - credito (lo que realmente impacta al conductor)

    Los KPIs del dashboard deben usar *_neto para no inflar los totales.
    """
    out = df.copy()

    # ── Renta: siempre negativa en Excel → abs() es correcto ──────────
    out["renta_semanal"] = out["renta_raw"].fillna(0).abs()

    # ── Fianza: invertir signo (neg=cobro, pos=devolución) ─────────────
    out["fianza"] = out["fianza_raw"].apply(_invertir_signo_fianza)

    # Desglose fianza para auditoría
    _fc, _fd, _fn = split_cargos_creditos(out["fianza_raw"])
    out["fianza_cobrada"]  = _fc   # cuánto se cobró al conductor
    out["fianza_devuelta"] = _fd   # cuánto se devolvió
    out["fianza_neta"]     = _fn   # cobrado - devuelto

    # ── Multa: signo mixto → NO usar abs() ────────────────────────────
    _mc, _mcr, _mn = split_cargos_creditos(out["multa_raw"])
    out["multa_cargo"]   = _mc    # cargos cobrados al conductor
    out["multa_credito"] = _mcr   # créditos/devoluciones
    out["multa_neto"]    = _mn    # neto real
    # Alias de compatibilidad (apunta al neto, no al abs)
    out["multa"]         = _mn

    # ── Hojalatero: signo mixto → NO usar abs() ───────────────────────
    _hc, _hcr, _hn = split_cargos_creditos(out["hojalatero_raw"])
    out["hojalatero_cargo"]   = _hc
    out["hojalatero_credito"] = _hcr
    out["hojalatero_neto"]    = _hn
    out["hojalatero"]         = _hn

    # ── Descuentos: signo mixto → el más crítico (muchos créditos) ────
    _dc, _dcr, _dn = split_cargos_creditos(out["descuentos_raw"])
    out["descuentos_cargo"]   = _dc
    out["descuentos_credito"] = _dcr
    out["descuentos_neto"]    = _dn
    out["descuentos"]         = _dn

    # ── Concepto auto-generado ─────────────────────────────────────────
    out["concepto_ingreso"] = out.apply(generar_concepto_ingreso, axis=1)

    # Mantener columnas raw para auditoría (no se eliminan)
    return out


# ═══════════════════════════════════════════════════════════════
# TABLA DE AUDITORÍA DE SIGNOS
# ═══════════════════════════════════════════════════════════════
def auditoria_signos(df_ing: pd.DataFrame) -> pd.DataFrame:
    """
    Genera tabla resumen para comparar el dashboard contra el Excel.

    Columnas:
      Concepto | Suma Abs | Cargos (neg) | Creditos (pos) | Neto | Filas neg | Filas pos | Filas nulas
    """
    CONCEPTOS = [
        ("Renta",       "renta_raw",       None,               None,               None),
        ("Fianza",      "fianza_raw",       "fianza_cobrada",   "fianza_devuelta",  "fianza_neta"),
        ("Multa",       "multa_raw",        "multa_cargo",      "multa_credito",    "multa_neto"),
        ("Hojalatero",  "hojalatero_raw",   "hojalatero_cargo", "hojalatero_credito", "hojalatero_neto"),
        ("Descuentos",  "descuentos_raw",   "descuentos_cargo", "descuentos_credito", "descuentos_neto"),
    ]

    filas = []
    for nombre, col_raw, col_cargo, col_cred, col_neto in CONCEPTOS:
        if col_raw not in df_ing.columns:
            continue
        s = df_ing[col_raw].fillna(0)
        negs  = (s < 0).sum()
        poss  = (s > 0).sum()
        nulas = df_ing[col_raw].isna().sum()

        suma_abs = s.abs().sum()

        if col_cargo and col_cargo in df_ing.columns:
            cargo  = df_ing[col_cargo].sum()
            cred   = df_ing[col_cred].sum()
            neto   = df_ing[col_neto].sum()
        else:
            # Renta: siempre negativa, abs() es correcto
            cargo  = s.abs().sum()
            cred   = 0.0
            neto   = cargo

        filas.append({
            "Concepto":        nombre,
            "Suma Abs":        suma_abs,
            "Cargos (neg)":    cargo,
            "Creditos (pos)":  cred,
            "Neto":            neto,
            "Filas negativas": negs,
            "Filas positivas": poss,
            "Filas nulas":     nulas,
        })

    return pd.DataFrame(filas)


# ═══════════════════════════════════════════════════════════════
# TRANSFORMACIONES — EGRESOS
# ═══════════════════════════════════════════════════════════════
def transform_egresos(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out["monto_real"] = out["monto_real"].fillna(0).abs()
    return out


# ═══════════════════════════════════════════════════════════════
# CARGA PRINCIPAL
# ═══════════════════════════════════════════════════════════════
def load_ingresos(
    path: str = "prueba.xlsx",
    años: list[int] | None = None,
) -> pd.DataFrame:
    """Lee hojas UBER 2025/2026 y devuelve DataFrame unificado y transformado."""
    if años is None:
        años = [2025, 2026]

    frames = []
    for año in años:
        cfg = SHEETS["ingresos"].get(año)
        if cfg is None:
            continue
        try:
            raw = pd.read_excel(path, sheet_name=cfg["name"], header=cfg["header_row"])
        except Exception:
            continue

        raw = _strip_unnamed(raw)
        raw.columns = [str(c).strip() for c in raw.columns]

        # Resolver aliases de columnas de forma robusta
        rename_map = _resolve_aliases(list(raw.columns), COLUMN_ALIASES)
        df = raw.rename(columns=rename_map)

        # Coerción numérica
        for col in _INGRESOS_NUMERIC:
            if col in df.columns:
                df[col] = _coerce_numeric(df[col])

        df["semana"] = _to_semana(df.get("semana"))
        df["año"] = año

        # Filtrar filas sin semana válida (totales, encabezados, etc.)
        df = df.dropna(subset=["semana"])

        frames.append(df)

    if not frames:
        return pd.DataFrame()

    combined = pd.concat(frames, ignore_index=True)
    return transform_ingresos(combined)


def load_egresos(
    path: str = "prueba.xlsx",
    años: list[int] | None = None,
) -> pd.DataFrame:
    """Lee hojas Gastos 2025/2026 y devuelve DataFrame unificado y transformado."""
    if años is None:
        años = [2025, 2026]

    frames = []
    for año in años:
        cfg = SHEETS["egresos"].get(año)
        if cfg is None:
            continue
        try:
            raw = pd.read_excel(path, sheet_name=cfg["name"], header=cfg["header_row"])
        except Exception:
            continue

        raw = _strip_unnamed(raw)
        raw.columns = [str(c).strip() for c in raw.columns]

        rename = _EGRESOS_RENAME.get(año, {})
        rename_valid = {k: v for k, v in rename.items() if k in raw.columns}
        df = raw.rename(columns=rename_valid)

        for col in _EGRESOS_NUMERIC:
            if col in df.columns:
                df[col] = _coerce_numeric(df[col])

        df["semana"] = _to_semana(df.get("semana"))
        df["año"] = año
        df = df.dropna(subset=["semana"])
        frames.append(df)

    if not frames:
        return pd.DataFrame()

    combined = pd.concat(frames, ignore_index=True)
    return transform_egresos(combined)


# ═══════════════════════════════════════════════════════════════
# UTILIDADES PARA DASHBOARD
# ═══════════════════════════════════════════════════════════════
def yearweek_key(año, semana) -> int:
    return int(año) * 100 + int(semana)


def yearweek_label(key: int) -> str:
    y = key // 100
    w = key % 100
    return f"{y}-S{w:02d}"


def add_yearweek(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    año_f    = pd.to_numeric(out["año"],    errors="coerce")
    semana_f = pd.to_numeric(out["semana"], errors="coerce")
    out["YEARWEEK"]   = año_f * 100 + semana_f
    out["WEEK_LABEL"] = (
        out["YEARWEEK"]
        .dropna()
        .astype(int)
        .map(yearweek_label)
    )
    return out


# ═══════════════════════════════════════════════════════════════
# CONSTANTES DE INVERSIÓN
# ═══════════════════════════════════════════════════════════════
CONCEPTO_SUBASTA = "PAGO DE SUBASTA"
PAGO_SEMANAL_DEFAULT = 2_500.0

CONCEPTOS_INVERSION: list[str] = [
    "INVERSIÓN",
    "INVERSION",
    "PAGO DE SUBASTA",
]


def _mask_inversion(df: pd.DataFrame) -> pd.Series:
    if "concepto" not in df.columns or df.empty:
        return pd.Series(False, index=df.index)
    conceptos_upper = [c.upper().strip() for c in CONCEPTOS_INVERSION]
    return (
        df["concepto"]
        .astype(str)
        .str.strip()
        .str.upper()
        .isin(conceptos_upper)
    )


def get_gastos(df_egresos: pd.DataFrame) -> pd.DataFrame:
    if df_egresos.empty:
        return df_egresos.copy()
    return df_egresos[~_mask_inversion(df_egresos)].copy()


def get_inversiones_egreso(df_egresos: pd.DataFrame) -> pd.DataFrame:
    if df_egresos.empty:
        return df_egresos.copy()
    return df_egresos[_mask_inversion(df_egresos)].copy()


def conductores_por_semana(df_ingresos: pd.DataFrame) -> pd.DataFrame:
    if df_ingresos.empty:
        return pd.DataFrame(columns=["YEARWEEK", "WEEK_LABEL", "n_conductores"])
    required = {"año", "semana", "conductor"}
    if required - set(df_ingresos.columns):
        return pd.DataFrame(columns=["YEARWEEK", "WEEK_LABEL", "n_conductores"])

    dedup = (
        df_ingresos
        .dropna(subset=["semana", "conductor"])
        .drop_duplicates(subset=["año", "semana", "conductor"])
    )
    if "YEARWEEK" not in dedup.columns:
        dedup = add_yearweek(dedup)

    return (
        dedup
        .groupby(["YEARWEEK", "WEEK_LABEL"], as_index=False)["conductor"]
        .nunique()
        .rename(columns={"conductor": "n_conductores"})
        .sort_values("YEARWEEK")
        .reset_index(drop=True)
    )


def calcular_amortizacion(
    capital: float,
    pago_semanal: float = PAGO_SEMANAL_DEFAULT,
) -> pd.DataFrame:
    if capital <= 0 or pago_semanal <= 0:
        return pd.DataFrame(columns=["semana_num", "pago_recuperacion", "saldo_restante"])
    filas = []
    saldo = float(capital)
    semana = 1
    while saldo > 0:
        pago = min(pago_semanal, saldo)
        saldo = round(saldo - pago, 2)
        filas.append({"semana_num": semana, "pago_recuperacion": pago, "saldo_restante": saldo})
        semana += 1
    return pd.DataFrame(filas)


def get_analisis_inversiones(
    df_egresos: pd.DataFrame,
    df_ingresos: pd.DataFrame,
    pago_semanal: float = PAGO_SEMANAL_DEFAULT,
) -> dict:
    resumen_filas = []
    tablas: dict[str, pd.DataFrame] = {}

    if df_egresos.empty or "concepto" not in df_egresos.columns:
        return {"resumen": pd.DataFrame(), "tablas": {}}

    subastas = df_egresos[
        df_egresos["concepto"].astype(str).str.strip().str.upper()
        == CONCEPTO_SUBASTA.upper()
    ].copy()

    if subastas.empty or "llave" not in subastas.columns:
        return {"resumen": pd.DataFrame(), "tablas": {}}

    llaves_subasta = [
        str(l) for l in subastas["llave"].dropna().unique()
        if str(l).strip() not in ("", "nan", "-")
    ]

    for llave in llaves_subasta:
        mask_sub = subastas["llave"].astype(str) == llave
        capital = float(subastas.loc[mask_sub, "monto_real"].fillna(0).sum())
        if capital <= 0:
            continue

        tabla = calcular_amortizacion(capital, pago_semanal)
        tablas[llave] = tabla

        semanas_recuperacion = int(len(tabla))
        monto_recuperado = float(tabla["pago_recuperacion"].sum())

        ingresos_veh = 0.0
        if not df_ingresos.empty and "llave" in df_ingresos.columns:
            mask_ing = df_ingresos["llave"].astype(str) == llave
            col_ing = "ganancias_totales" if "ganancias_totales" in df_ingresos.columns else "renta_semanal"
            ingresos_veh = float(df_ingresos.loc[mask_ing, col_ing].fillna(0).sum())

        roi_pct = ((ingresos_veh - capital) / capital * 100) if capital > 0 else 0.0

        semanas_activas = 0
        if not df_ingresos.empty and "llave" in df_ingresos.columns:
            mask_ing = df_ingresos["llave"].astype(str) == llave
            semanas_activas = int(df_ingresos.loc[mask_ing, "semana"].dropna().nunique())

        resumen_filas.append({
            "llave":                llave,
            "capital_inicial":      capital,
            "pago_semanal":         pago_semanal,
            "semanas_recuperacion": semanas_recuperacion,
            "monto_recuperado":     monto_recuperado,
            "ingresos_totales":     ingresos_veh,
            "semanas_activas":      semanas_activas,
            "roi_pct":              round(roi_pct, 2),
            "recuperado":           monto_recuperado >= capital,
        })

    resumen = (
        pd.DataFrame(resumen_filas)
        .sort_values("capital_inicial", ascending=False)
        .reset_index(drop=True)
    )
    return {"resumen": resumen, "tablas": tablas}


# ═══════════════════════════════════════════════════════════════
# EJECUCIÓN DIRECTA (prueba rápida)
# ═══════════════════════════════════════════════════════════════
if __name__ == "__main__":
    import sys, io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

    print("=" * 60)
    print("Cargando prueba.xlsx ...")
    print("=" * 60)

    df_ing = load_ingresos()
    print(f"\n Ingresos: {len(df_ing)} filas, anios: {sorted(df_ing['anio'].unique()) if 'anio' in df_ing.columns else sorted(df_ing['año'].unique())}")

    # Prueba de signos 2026
    df26 = df_ing[df_ing["año"] == 2026]
    print("\n--- VALIDACION SIGNOS 2026 ---")
    for concepto, col_raw, col_cargo, col_cred, col_neto in [
        ("Multa",      "multa_raw",       "multa_cargo",       "multa_credito",       "multa_neto"),
        ("Hojalatero", "hojalatero_raw",  "hojalatero_cargo",  "hojalatero_credito",  "hojalatero_neto"),
        ("Descuentos", "descuentos_raw",  "descuentos_cargo",  "descuentos_credito",  "descuentos_neto"),
    ]:
        if col_raw not in df26.columns:
            print(f"  {concepto}: columna raw no encontrada")
            continue
        print(f"\n  {concepto}:")
        print(f"    raw sum         : {df26[col_raw].sum():>12,.2f}")
        print(f"    cargo (cargos)  : {df26[col_cargo].sum():>12,.2f}")
        print(f"    credito         : {df26[col_cred].sum():>12,.2f}")
        print(f"    neto (KPI nuevo): {df26[col_neto].sum():>12,.2f}")
        print(f"    abs() (anterior): {df26[col_raw].fillna(0).abs().sum():>12,.2f}")

    # Tabla de auditoria
    print("\n--- TABLA DE AUDITORIA ---")
    aud = auditoria_signos(df26)
    print(aud.to_string(index=False))

    print("\n business_rules.py OK")
