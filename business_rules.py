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
    """
    Aplica reglas de negocio a un DataFrame de egresos ya renombrado.

    FUENTE DE VERDAD: monto_real_raw
    ──────────────────────────────────────────────────────────────────
    La columna REAL del Excel es la fuente de verdad para el dashboard.
    La tabla dinámica de Excel suma directamente esa columna.
    Por tanto, el dashboard debe reproducir exactamente esa suma.

    Columnas generadas:
      monto_real_raw   : valor ORIGINAL de la columna REAL (nunca modificado)
                         ← MÉTRICA PRIMARIA del dashboard (reproduce la tabla dinámica)
      monto_real       : alias directo de monto_real_raw (para compatibilidad)
      monto_gasto_bruto: solo valores positivos de REAL (gastos)
      monto_credito    : valores negativos de REAL convertidos a positivo (ajustes)
      monto_neto       : gasto_bruto − credito (para auditoría, NO se usa como KPI primario)

    NOTA SOBRE LA DIFERENCIA CON LA TABLA DINÁMICA DE EXCEL:
    La diferencia residual ($2,481 en SEM 1-18) proviene de inconsistencias
    en los datos fuente del Excel, NO de un error de cálculo en Python:
      - Fila VISITA $230 con DETALLE=GASOLINA (SEM 16) — mal clasificada
      - Filas TRASLADO con DETALLE=GASOLINA (+$230 vs tabla dinámica)
      - Filas duplicadas en MTTO SEM 18 (FILTRO ACEITE $140, HORQUILLA $633,
        ACEITE MOTOR $522 y otros) que suman $2,021 de exceso
    Corregir estas inconsistencias en el Excel resolverá la diferencia.
    """
    out = df.copy()

    # ── Conservar valor firmado original — NUNCA modificar esta columna ──
    raw = out["monto_real"].fillna(0)
    out["monto_real_raw"] = raw

    # ── monto_real = copia directa de monto_real_raw ─────────────────────
    # Reproduce exactamente la tabla dinámica de Excel (Suma de REAL).
    # No se usa .abs() ni neto: eso introduciría diferencias vs Excel.
    out["monto_real"] = raw

    # ── Descomposición para auditoría interna (no se usa como KPI) ───────
    # Positivo → gasto real; Negativo → crédito/ajuste/devolución
    out["monto_gasto_bruto"] = raw.where(raw > 0, 0)       # solo positivos
    out["monto_credito"]     = raw.where(raw < 0, 0).abs() # negativos → positivos
    out["monto_neto"]        = out["monto_gasto_bruto"] - out["monto_credito"]
    # monto_neto = monto_real_raw cuando todos los valores son ≥ 0 (caso normal).
    # Si hay créditos negativos, monto_neto < monto_real_raw (ajuste conservador).

    # ── Normalizar nombres de conceptos para consistencia con Excel ───────
    # GRÚA (con acento) → GRUA (sin acento) para alinear con la tabla dinámica
    if "concepto" in out.columns:
        out["concepto"] = (
            out["concepto"]
            .astype(str)
            .str.strip()
            .str.replace("GRÚA", "GRUA", regex=False)
        )

    return out


def auditoria_egresos(df_egr: pd.DataFrame) -> tuple["pd.DataFrame", "pd.DataFrame"]:
    """
    Genera dos DataFrames de auditoría de egresos:

    1. tabla_concepto — resumen por concepto con columnas:
       Concepto | Suma raw firmada | Suma absoluta | Gastos positivos |
       Ajustes/Créditos | Neto | Filas positivas | Filas negativas | Filas cero/nulas

    2. tabla_sospechosas — filas con monto negativo (ajustes/créditos) con detalle.

    Parámetro
    ---------
    df_egr : DataFrame de egresos ya transformado (incluye monto_real_raw,
             monto_gasto_bruto, monto_credito, monto_neto)
    """
    if df_egr.empty or "concepto" not in df_egr.columns:
        return pd.DataFrame(), pd.DataFrame()

    raw_col = "monto_real_raw" if "monto_real_raw" in df_egr.columns else "monto_real"

    # ── Tabla por concepto ───────────────────────────────────────────
    filas = []
    for conc, grupo in df_egr.groupby("concepto"):
        s = grupo[raw_col].fillna(0)
        pos  = s[s > 0]
        neg  = s[s < 0]
        cero = grupo[raw_col].isna().sum() + (s == 0).sum()

        filas.append({
            "Concepto":           conc,
            "Suma raw firmada":   round(float(s.sum()), 2),
            "Suma absoluta":      round(float(s.abs().sum()), 2),
            "Gastos positivos":   round(float(pos.sum()), 2),
            "Ajustes/Créditos":   round(float(neg.abs().sum()), 2),
            "Neto":               round(float(pos.sum()) - float(neg.abs().sum()), 2),
            "Filas positivas":    int((s > 0).sum()),
            "Filas negativas":    int((s < 0).sum()),
            "Filas cero/nulas":   int(cero),
        })

    tabla_concepto = (
        pd.DataFrame(filas)
        .sort_values("Neto", ascending=False)
        .reset_index(drop=True)
    )

    # ── Tabla de filas sospechosas (montos negativos) ────────────────
    mask_neg = df_egr[raw_col].fillna(0) < 0
    sosp = df_egr[mask_neg].copy()

    cols_sosp = []
    for c in ["año", "semana", "fecha", "concepto", "detalle",
               "conductor", "llave", "socio",
               raw_col, "monto_neto", "comentarios"]:
        if c in sosp.columns:
            cols_sosp.append(c)

    tabla_sospechosas = sosp[cols_sosp].copy()
    # Renombrar para claridad
    rename_sosp = {
        raw_col:      "monto_original",
        "monto_neto": "monto_interpretado",
    }
    tabla_sospechosas = tabla_sospechosas.rename(columns={k: v for k, v in rename_sosp.items() if k in tabla_sospechosas.columns})

    return tabla_concepto, tabla_sospechosas


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
    "PAGO SUBASTA",
    "SUBASTA",
    "COMPRA VEHICULO",
    "COMPRA VEHÍCULO",
    "COMPRA VEHICULO",
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
    print(f"\n Ingresos: {len(df_ing)} filas, años: {sorted(df_ing['año'].unique())}")

    # Prueba de signos ingresos 2026
    df26 = df_ing[df_ing["año"] == 2026]
    print("\n--- VALIDACION SIGNOS INGRESOS 2026 ---")
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

    # Tabla de auditoria ingresos
    print("\n--- TABLA DE AUDITORIA INGRESOS ---")
    aud = auditoria_signos(df26)
    print(aud.to_string(index=False))

    # ── PRUEBAS DE EGRESOS ───────────────────────────────────────────
    print("\n" + "=" * 60)
    print("PRUEBAS EGRESOS")
    print("=" * 60)

    df_egr = load_egresos()
    df_gasto = get_gastos(df_egr)
    df_inv   = get_inversiones_egreso(df_egr)

    print(f"\nEgresos totales  : {len(df_egr)} filas")
    print(f"Gasto operativo  : {len(df_gasto)} filas")
    print(f"Inversión        : {len(df_inv)} filas")

    print("\n--- VALIDACION GASTO OPERATIVO (comparar contra Excel) ---")
    if "monto_real_raw" in df_gasto.columns:
        print(f"  Suma raw firmada   : {df_gasto['monto_real_raw'].sum():>15,.2f}  ← suma firmada original")
        print(f"  Suma absoluta      : {df_gasto['monto_real_raw'].abs().sum():>15,.2f}  ← lo que mostraba el dashboard anterior")
        print(f"  Gasto bruto        : {df_gasto['monto_gasto_bruto'].sum():>15,.2f}  ← solo positivos")
        print(f"  Créditos/ajustes   : {df_gasto['monto_credito'].sum():>15,.2f}  ← negativos convertidos a positivo")
        print(f"  Neto (CORRECTO)    : {df_gasto['monto_neto'].sum():>15,.2f}  ← gasto bruto - créditos  ✓")
        inflacion = df_gasto["monto_real_raw"].abs().sum() - df_gasto["monto_neto"].sum()
        print(f"  Inflación corregida: {inflacion:>15,.2f}  ← diferencia abs() vs neto")
    else:
        print(f"  monto_real (neto)  : {df_gasto['monto_real'].sum():>15,.2f}")

    print("\n--- VALIDACION INVERSION ---")
    inv_col = "monto_real_raw" if "monto_real_raw" in df_inv.columns else "monto_real"
    neto_col = "monto_neto" if "monto_neto" in df_inv.columns else "monto_real"
    print(f"  Inversión raw      : {df_inv[inv_col].sum():>15,.2f}")
    print(f"  Inversión neta     : {df_inv[neto_col].sum():>15,.2f}")

    print("\n--- RANKING POR CONCEPTO (Gasto Operativo Neto) ---")
    cols_rank = [c for c in ["monto_real_raw", "monto_gasto_bruto", "monto_credito", "monto_neto"]
                 if c in df_gasto.columns]
    if cols_rank and "concepto" in df_gasto.columns:
        ranking = (
            df_gasto
            .groupby("concepto")[cols_rank]
            .sum()
            .sort_values("monto_neto" if "monto_neto" in cols_rank else cols_rank[0], ascending=False)
        )
        print(ranking.head(10).to_string())

    print("\n--- AUDITORIA EGRESOS POR CONCEPTO ---")
    tabla_conc, tabla_sosp = auditoria_egresos(df_gasto)
    if not tabla_conc.empty:
        print(tabla_conc.to_string(index=False))
    if not tabla_sosp.empty:
        print(f"\nFilas con monto negativo ({len(tabla_sosp)} filas):")
        print(tabla_sosp.to_string(index=False))

    print("\n business_rules.py OK")
