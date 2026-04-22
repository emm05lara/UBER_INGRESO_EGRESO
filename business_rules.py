# business_rules.py
# ──────────────────────────────────────────────────────────────
# Módulo de reglas de negocio — Operación UBER 2025 / 2026
# Lee prueba.xlsx (4 hojas) y expone DataFrames limpios.
# ──────────────────────────────────────────────────────────────
from __future__ import annotations

import numpy as np
import pandas as pd

# ═══════════════════════════════════════════════════════════════
# CONFIGURACIÓN DE HOJAS
# ═══════════════════════════════════════════════════════════════
SHEETS = {
    "ingresos": {
        2025: {"name": "UBER 2025", "header_row": 0},   # fila 0 = encabezados reales
        2026: {"name": "UBER 2026", "header_row": 0},
    },
    "egresos": {
        2025: {"name": "Gastos 2025", "header_row": 0},
        2026: {"name": "Gastos 2026", "header_row": 0},
    },
}

# ═══════════════════════════════════════════════════════════════
# MAPEO DE COLUMNAS — INGRESOS  (UBER 2025 / 2026)
# ═══════════════════════════════════════════════════════════════
# Nombres canónicos → nombres reales por año
_INGRESOS_RENAME = {
    2025: {
        # ── Columnas reales en UBER 2025 (header_row=0) ──────────────────
        "SEM":             "semana",
        "CONDUCTOR":       "conductor",
        "AUTO":            "llave",        # 2025 usa AUTO, 2026 usa LLAVE
        "TAG":             "tag",
        "SOCIO":           "socio",
        "PLATAFORMA":      "plataforma",
        "APP":             "app",
        "GANANCIA":        "ganancia",
        "RENTA":           "renta_raw",
        "SEM PASADA":      "sem_pasada",
        "FIANZA":          "fianza_raw",
        "MULTA":           "multa_raw",
        "HOJALATERO":      "hojalatero_raw",   # Excel 2025 tiene nombre largo
        "DESCUENTOS":      "descuentos_raw",
        "TOTAL":           "total",
        "GANANCIAS TOTALES": "ganancias_totales",
        "COMENTARIOS":     "comentarios",
    },
    2026: {
        # ── Columnas reales en UBER 2026 (header_row=0) ──────────────────
        "SEM":       "semana",
        "CONDUCTOR": "conductor",
        "LLAVE":     "llave",            # 2026 usa LLAVE, 2025 usa AUTO
        "TAG":       "tag",
        "SOCIO":     "socio",
        "PLATAFORMA":"plataforma",
        "APP":       "app",
        "GANANCIA":  "ganancia",
        "RENTA":     "renta_raw",
        "SEM A":     "sem_pasada",       # ← era "SEM PASADA", ahora abreviado
        "FIANZA":    "fianza_raw",
        "MULTA":     "multa_raw",
        "HOJALAT":   "hojalatero_raw",   # ← era "HOJALATERO", ahora abreviado
        "DESC":      "descuentos_raw",   # ← era "DESCUENTOS", ahora abreviado
        "TOTAL":     "total",
        "GANAN T":   "ganancias_totales", # ← era "GANANCIAS \nTOTALES"
        "COMENTARIOS": "comentarios",
    },
}

# ═══════════════════════════════════════════════════════════════
# MAPEO DE COLUMNAS — EGRESOS  (Gastos 2025 / 2026)
# ═══════════════════════════════════════════════════════════════
_EGRESOS_RENAME = {
    2025: {
        "Semana": "semana",
        "MES": "mes",
        "Fecha": "fecha",
        "Solicitante": "solicitante",
        "CONCEPTO": "concepto",
        "DETALLE": "detalle",
        "CONDUCTOR": "conductor",
        "DETALLE.1": "llave",            # 2025 usa segunda col DETALLE como llave
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

# Columnas numéricas que siempre se convierten
_INGRESOS_NUMERIC = [
    "ganancia", "renta_raw", "sem_pasada", "fianza_raw",
    "multa_raw", "hojalatero_raw", "descuentos_raw",
    "total", "ganancias_totales",
]

_EGRESOS_NUMERIC = ["monto_real"]


# ═══════════════════════════════════════════════════════════════
# HELPERS
# ═══════════════════════════════════════════════════════════════
def _to_semana(s: pd.Series) -> pd.Series:
    """Convierte la columna de semana a float64 (compatible con NaN y PyArrow).

    Se usa float64 en lugar de Int64 nullable de pandas porque este último
    no es reconocido por PyArrow (motor de serialización de Streamlit) en
    ciertas versiones, causando el error: data type 'Int64' not understood.
    """
    return pd.to_numeric(s, errors="coerce")  # devuelve float64 con NaN donde no parsea


def _coerce_numeric(s: pd.Series) -> pd.Series:
    """Convierte a numérico aceptando $, comas, paréntesis (negativos)."""
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
        if "." in val:
            return val.replace(",", "")
        return val.replace(",", ".")

    x = x.apply(_fix_commas)
    return pd.to_numeric(x, errors="coerce")


def _strip_unnamed(df: pd.DataFrame) -> pd.DataFrame:
    return df.loc[:, ~df.columns.astype(str).str.match(r"^Unnamed")]


# ═══════════════════════════════════════════════════════════════
# TRANSFORMACIONES — INGRESOS
# ═══════════════════════════════════════════════════════════════
def _invertir_signo_fianza(val):
    """Regla: en el Excel '-' representa ingreso de dinero,
    '+' es devolución al conductor.  Invertimos para que
    positivo = ingreso, negativo = devolución."""
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
    if abs(row.get("multa", 0) or 0) > 0:
        partes.append("Multa")
    if abs(row.get("hojalatero", 0) or 0) > 0:
        partes.append("Hojalatero")
    if abs(row.get("descuentos", 0) or 0) > 0:
        partes.append("Descuento")
    return " + ".join(partes) if partes else "Sin concepto"


def transform_ingresos(df: pd.DataFrame) -> pd.DataFrame:
    """Aplica reglas de negocio a un DataFrame de ingresos ya renombrado."""
    out = df.copy()

    # Valor absoluto: renta, multa, hojalatero, descuentos
    out["renta_semanal"] = out["renta_raw"].fillna(0).abs()
    out["multa"]         = out["multa_raw"].fillna(0).abs()
    out["hojalatero"]    = out["hojalatero_raw"].fillna(0).abs()
    out["descuentos"]    = out["descuentos_raw"].fillna(0).abs()

    # Fianza: invertir signo
    out["fianza"] = out["fianza_raw"].apply(_invertir_signo_fianza)

    # Concepto auto-generado
    out["concepto_ingreso"] = out.apply(generar_concepto_ingreso, axis=1)

    # Limpiar columnas raw
    out.drop(columns=[
        "renta_raw", "multa_raw", "hojalatero_raw",
        "descuentos_raw", "fianza_raw",
    ], inplace=True)

    return out


# ═══════════════════════════════════════════════════════════════
# TRANSFORMACIONES — EGRESOS
# ═══════════════════════════════════════════════════════════════
def transform_egresos(df: pd.DataFrame) -> pd.DataFrame:
    """Aplica reglas de negocio a un DataFrame de egresos ya renombrado."""
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
    """Lee hojas UBER 2025 / 2026 y devuelve DataFrame unificado y transformado."""
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

        rename = _INGRESOS_RENAME.get(año, {})
        # Solo renombrar columnas que existen
        rename_valid = {k: v for k, v in rename.items() if k in raw.columns}
        df = raw.rename(columns=rename_valid)

        # Coerción numérica
        for col in _INGRESOS_NUMERIC:
            if col in df.columns:
                df[col] = _coerce_numeric(df[col])

        df["semana"] = _to_semana(df.get("semana"))  # float64, acepta NaN, compatible PyArrow
        df["año"] = año

        # Filtrar filas sin semana (encabezados, totales)
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
    """Lee hojas Gastos 2025 / 2026 y devuelve DataFrame unificado y transformado."""
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

        df["semana"] = _to_semana(df.get("semana"))  # float64, acepta NaN, compatible PyArrow
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
    """Crea un entero AÑO*100 + SEM para ordenar."""
    return int(año) * 100 + int(semana)


def yearweek_label(key: int) -> str:
    y = key // 100
    w = key % 100
    return f"{y}-S{w:02d}"


def add_yearweek(df: pd.DataFrame) -> pd.DataFrame:
    """Agrega columnas YEARWEEK y WEEK_LABEL para agrupar por semana."""
    out = df.copy()
    # Usamos float64 para evitar el nullable Int64 de pandas, incompatible con PyArrow.
    # Las filas sin semana válida quedarán como NaN y se filtran en el dashboard.
    año_f   = pd.to_numeric(out["año"],    errors="coerce")  # float64
    semana_f = pd.to_numeric(out["semana"], errors="coerce")  # float64
    out["YEARWEEK"] = año_f * 100 + semana_f                 # float64, NaN-safe
    out["WEEK_LABEL"] = (
        out["YEARWEEK"]
        .dropna()
        .astype(int)          # solo para construir la etiqueta string
        .map(yearweek_label)
    )
    return out


# ═══════════════════════════════════════════════════════════════
# CONSTANTES DE INVERSIÓN
# ═══════════════════════════════════════════════════════════════
CONCEPTO_SUBASTA = "PAGO DE SUBASTA"
PAGO_SEMANAL_DEFAULT = 2_500.0  # MXN por semana

# Conceptos que se clasifican como INVERSIÓN (no gasto operativo).
# Case-insensitive. Si detectas variantes en el Excel, repórtalas.
CONCEPTOS_INVERSION: list[str] = [
    "INVERSIÓN",
    "INVERSION",   # sin acento, por si acaso
    "PAGO DE SUBASTA",
]


# ═══════════════════════════════════════════════════════════════
# UTILIDADES DE FILTRADO DE EGRESOS
# ═══════════════════════════════════════════════════════════════
def _mask_inversion(df: pd.DataFrame) -> pd.Series:
    """Devuelve máscara booleana True donde el concepto es INVERSIÓN.

    Compara en mayúsculas y sin espacios extra para ser robusto
    ante variaciones de capitalización en el Excel.
    """
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
    """Filtra el DataFrame de egresos excluyendo conceptos de INVERSIÓN.

    Usar esta función en todos los KPIs y gráficas de «gasto operativo»
    para que la inversión no distorsione los totales.

    Returns
    -------
    pd.DataFrame — egresos sin filas de inversión.
    """
    if df_egresos.empty:
        return df_egresos.copy()
    return df_egresos[~_mask_inversion(df_egresos)].copy()


def get_inversiones_egreso(df_egresos: pd.DataFrame) -> pd.DataFrame:
    """Devuelve solo las filas de egresos clasificadas como INVERSIÓN.

    Usar esta función para mostrar el KPI de inversión separado del gasto.

    Returns
    -------
    pd.DataFrame — únicamente los egresos de tipo inversión.
    """
    if df_egresos.empty:
        return df_egresos.copy()
    return df_egresos[_mask_inversion(df_egresos)].copy()


def conductores_por_semana(df_ingresos: pd.DataFrame) -> pd.DataFrame:
    """Cuenta conductores únicos por semana, eliminando duplicados.

    Puede haber filas repetidas en el Excel (mismo conductor, misma semana).
    Esta función deduplica por (año, semana, conductor) antes de contar,
    evitando el sobreconteo.

    Returns
    -------
    pd.DataFrame con columnas:
        YEARWEEK       – clave numérica año*100+sem
        WEEK_LABEL     – etiqueta «YYYY-SNN»
        n_conductores  – número de conductores únicos esa semana
    """
    if df_ingresos.empty:
        return pd.DataFrame(
            columns=["YEARWEEK", "WEEK_LABEL", "n_conductores"]
        )

    required = {"año", "semana", "conductor"}
    missing = required - set(df_ingresos.columns)
    if missing:
        return pd.DataFrame(
            columns=["YEARWEEK", "WEEK_LABEL", "n_conductores"]
        )

    # Deduplicar: un conductor cuenta una sola vez por semana
    dedup = (
        df_ingresos
        .dropna(subset=["semana", "conductor"])
        .drop_duplicates(subset=["año", "semana", "conductor"])
    )

    # Asegurarnos de tener YEARWEEK y WEEK_LABEL
    if "YEARWEEK" not in dedup.columns:
        dedup = add_yearweek(dedup)

    # Contar conductores únicos por semana
    result = (
        dedup
        .groupby(["YEARWEEK", "WEEK_LABEL"], as_index=False)["conductor"]
        .nunique()
        .rename(columns={"conductor": "n_conductores"})
        .sort_values("YEARWEEK")
        .reset_index(drop=True)
    )
    return result


def calcular_amortizacion(
    capital: float,
    pago_semanal: float = PAGO_SEMANAL_DEFAULT,
) -> pd.DataFrame:
    """Genera la tabla de amortización semanal para un capital dado.

    Parameters
    ----------
    capital : float
        Monto invertido (capital inicial).
    pago_semanal : float
        Pago fijo aplicado por semana contra el saldo.

    Returns
    -------
    pd.DataFrame con columnas:
        semana_num          – número de semana (1, 2, 3, …)
        pago_recuperacion   – pago aplicado esa semana
        saldo_restante      – saldo pendiente después del pago
    """
    if capital <= 0 or pago_semanal <= 0:
        return pd.DataFrame(
            columns=["semana_num", "pago_recuperacion", "saldo_restante"]
        )

    filas = []
    saldo = float(capital)
    semana = 1
    while saldo > 0:
        pago = min(pago_semanal, saldo)
        saldo = round(saldo - pago, 2)
        filas.append({
            "semana_num": semana,
            "pago_recuperacion": pago,
            "saldo_restante": saldo,
        })
        semana += 1

    return pd.DataFrame(filas)


def get_analisis_inversiones(
    df_egresos: pd.DataFrame,
    df_ingresos: pd.DataFrame,
    pago_semanal: float = PAGO_SEMANAL_DEFAULT,
) -> dict:
    """Analiza la recuperación de inversión para vehículos con PAGO DE SUBASTA.

    Parameters
    ----------
    df_egresos : pd.DataFrame
        DataFrame completo de egresos (post-transformación).
    df_ingresos : pd.DataFrame
        DataFrame completo de ingresos (post-transformación).
    pago_semanal : float
        Monto semanal destinado a recuperar la inversión (default $2,500).

    Returns
    -------
    dict con dos claves:
        "resumen"  → DataFrame con una fila por vehículo y métricas agregadas.
        "tablas"   → dict { llave: DataFrame de amortización }.
    """
    import math

    resumen_filas = []
    tablas: dict[str, pd.DataFrame] = {}

    # -- Detectar subastas --------------------------------------------------------
    if df_egresos.empty or "concepto" not in df_egresos.columns:
        return {"resumen": pd.DataFrame(), "tablas": {}}

    subastas = df_egresos[
        df_egresos["concepto"].astype(str).str.strip().str.upper()
        == CONCEPTO_SUBASTA.upper()
    ].copy()

    if subastas.empty:
        return {"resumen": pd.DataFrame(), "tablas": {}}

    if "llave" not in subastas.columns:
        return {"resumen": pd.DataFrame(), "tablas": {}}

    llaves_subasta = [
        str(l) for l in subastas["llave"].dropna().unique() if str(l).strip() not in ("", "nan", "-")
    ]

    for llave in llaves_subasta:
        # Capital total invertido en este vehículo
        mask_sub = subastas["llave"].astype(str) == llave
        capital = float(subastas.loc[mask_sub, "monto_real"].fillna(0).sum())
        if capital <= 0:
            continue

        # Tabla de amortización
        tabla = calcular_amortizacion(capital, pago_semanal)
        tablas[llave] = tabla

        semanas_recuperacion = int(len(tabla))
        monto_recuperado = float(tabla["pago_recuperacion"].sum())

        # Ingresos generados por el vehículo (ganancias_totales)
        ingresos_veh = 0.0
        if not df_ingresos.empty and "llave" in df_ingresos.columns:
            mask_ing = df_ingresos["llave"].astype(str) == llave
            col_ing = "ganancias_totales" if "ganancias_totales" in df_ingresos.columns else "renta_semanal"
            ingresos_veh = float(df_ingresos.loc[mask_ing, col_ing].fillna(0).sum())

        # ROI: (ingresos_totales - capital_inicial) / capital_inicial * 100
        roi_pct = ((ingresos_veh - capital) / capital * 100) if capital > 0 else 0.0

        # Número de semanas activas en ingresos
        semanas_activas = 0
        if not df_ingresos.empty and "llave" in df_ingresos.columns:
            mask_ing = df_ingresos["llave"].astype(str) == llave
            semanas_activas = int(df_ingresos.loc[mask_ing, "semana"].dropna().nunique())

        resumen_filas.append({
            "llave": llave,
            "capital_inicial": capital,
            "pago_semanal": pago_semanal,
            "semanas_recuperacion": semanas_recuperacion,
            "monto_recuperado": monto_recuperado,
            "ingresos_totales": ingresos_veh,
            "semanas_activas": semanas_activas,
            "roi_pct": round(roi_pct, 2),
            "recuperado": monto_recuperado >= capital,
        })

    resumen = pd.DataFrame(resumen_filas).sort_values("capital_inicial", ascending=False).reset_index(drop=True)
    return {"resumen": resumen, "tablas": tablas}


# ═══════════════════════════════════════════════════════════════
# EJECUCIÓN DIRECTA (prueba rápida)
# ═══════════════════════════════════════════════════════════════
if __name__ == "__main__":
    print("=" * 60)
    print("Cargando prueba.xlsx ...")
    print("=" * 60)

    df_ing = load_ingresos()
    print(f"\n✅ Ingresos: {len(df_ing)} filas")
    print(f"   Columnas: {list(df_ing.columns)}")
    print(f"   Años: {sorted(df_ing['año'].unique())}")
    if len(df_ing) > 0:
        print(f"   Renta semanal (ejemplo): {df_ing['renta_semanal'].head(3).tolist()}")
        print(f"   Fianza (ejemplo):        {df_ing['fianza'].head(3).tolist()}")
        print(f"   Conceptos:               {df_ing['concepto_ingreso'].value_counts().head(5).to_dict()}")
        print(f"   Ganancias totales sum:    ${df_ing['ganancias_totales'].sum():,.2f}")

    df_egr = load_egresos()
    print(f"\n✅ Egresos: {len(df_egr)} filas")
    print(f"   Columnas: {list(df_egr.columns)}")
    print(f"   Años: {sorted(df_egr['año'].unique())}")
    if len(df_egr) > 0:
        print(f"   Monto real (ejemplo):     {df_egr['monto_real'].head(3).tolist()}")
        print(f"   Gasto total sum (CON inv):${df_egr['monto_real'].sum():,.2f}")
        if "concepto" in df_egr.columns:
            print(f"   Top conceptos:            {df_egr['concepto'].value_counts().head(8).to_dict()}")

    # ── Validación: separación inversión / gasto ──────────────────────────
    print("\n" + "─" * 60)
    print("VALIDACIÓN — Separación INVERSIÓN / GASTO OPERATIVO")
    print("─" * 60)
    df_gasto = get_gastos(df_egr)
    df_inv   = get_inversiones_egreso(df_egr)

    total_egr_bruto = df_egr["monto_real"].sum()   if len(df_egr)   > 0 else 0.0
    total_gasto_val = df_gasto["monto_real"].sum() if len(df_gasto) > 0 else 0.0
    total_inv_val   = df_inv["monto_real"].sum()   if len(df_inv)   > 0 else 0.0

    print(f"  Egresos TOTALES (con inversión):  ${total_egr_bruto:>12,.2f}")
    print(f"  Gasto OPERATIVO  (sin inversión): ${total_gasto_val:>12,.2f}")
    print(f"  INVERSIÓN separada:               ${total_inv_val:>12,.2f}")
    cuadre_ok = abs((total_gasto_val + total_inv_val) - total_egr_bruto) < 0.01
    print(f"  ✓ Cuadre (gasto + inv = total):   {'OK' if cuadre_ok else '⚠️ DIFERENCIA'}")

    if len(df_inv) > 0 and "concepto" in df_inv.columns:
        print(f"\n  Conceptos detectados como INVERSIÓN:")
        for concepto, cnt in df_inv["concepto"].value_counts().items():
            monto = df_inv.loc[df_inv["concepto"] == concepto, "monto_real"].sum()
            print(f"    • {str(concepto):30s}  {cnt:3d} filas  ${monto:,.2f}")
    else:
        print("  (No se encontraron filas de inversión en el período)")

    # ── Validación: conductores únicos por semana ─────────────────────────
    print("\n" + "─" * 60)
    print("VALIDACIÓN — Conductores únicos por semana")
    print("─" * 60)
    df_ing_yw = add_yearweek(df_ing) if "YEARWEEK" not in df_ing.columns else df_ing
    cond_sem = conductores_por_semana(df_ing_yw)
    if len(cond_sem) > 0:
        print(f"  Semanas con datos: {len(cond_sem)}")
        print(f"  Conductores únicos promedio/semana: {cond_sem['n_conductores'].mean():.1f}")
        print(f"  Máx: {cond_sem['n_conductores'].max()}  Mín: {cond_sem['n_conductores'].min()}")
        print(f"\n  Primeras filas:")
        print(cond_sem.head(5).to_string(index=False))
    else:
        print("  (Sin datos de conductores)")

    print("\n✅ Módulo business_rules.py OK")
