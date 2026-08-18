import streamlit as st
import pandas as pd
import numpy as np
import re
import unicodedata
from io import BytesIO
import plotly.graph_objects as go

# ==============================================================================
# CONFIGURACIÓN
# ==============================================================================
APP_VERSION = "3.0"
UMBRAL_TOLERANCIA = 1.0
UMBRAL_FOLIO = 0.01

# Prefijos documentales que sí tratamos como folios.
# Se conservan; NO se eliminan durante la normalización.
PREFIJOS_FOLIO = (
    "NCTA", "NCNCT", "SNCTA", "ANCT", "PNCT", "NTCA", "NTA", "NC"
)

MESES_ES = {
    "ene": 1, "feb": 2, "mar": 3, "abr": 4, "may": 5, "jun": 6,
    "jul": 7, "ago": 8, "sep": 9, "sept": 9, "oct": 10, "nov": 11, "dic": 12,
}

# ==============================================================================
# 1. UTILIDADES
# ==============================================================================

def quitar_acentos(texto):
    texto = "" if texto is None else str(texto)
    return "".join(
        c for c in unicodedata.normalize("NFKD", texto)
        if not unicodedata.combining(c)
    )


def texto_norm(texto):
    if pd.isna(texto):
        return ""
    s = quitar_acentos(str(texto)).upper().strip()
    s = re.sub(r"\s+", " ", s)
    return s


def concepto_norm(texto):
    """Normalización conservadora para comparar conceptos entre cuentas."""
    s = texto_norm(texto)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def parse_spanish_date(valor):
    if pd.isna(valor):
        return pd.NaT

    # Si Excel ya entregó Timestamp/datetime.
    if isinstance(valor, (pd.Timestamp, np.datetime64)):
        return pd.Timestamp(valor)

    s = str(valor).strip()

    # dd/Mmm/aaaa
    m = re.match(
        r"^(\d{1,2})[/\-]([A-Za-zÁÉÍÓÚáéíóúÑñ]{3,4})[/\-](\d{4})$",
        s,
    )
    if m:
        day, mon_abbr, year = m.groups()
        mon = quitar_acentos(mon_abbr).lower()
        if mon in MESES_ES:
            try:
                return pd.Timestamp(
                    year=int(year), month=MESES_ES[mon], day=int(day)
                )
            except ValueError:
                return pd.NaT

    # Fallback para fechas numéricas o formatos de Excel ya convertidos a texto.
    return pd.to_datetime(s, dayfirst=True, errors="coerce")


def parse_amount(valor):
    """
    Convierte montos de CONTPAQ sin convertir silenciosamente texto inválido en cero.
    Soporta:
      1,234.56
      $1,234.56
      (1,234.56)
      -1,234.56
      vacío -> 0
    """
    if pd.isna(valor):
        return 0.0

    if isinstance(valor, (int, float, np.integer, np.floating)):
        if pd.isna(valor):
            return 0.0
        return float(valor)

    s = str(valor).strip()
    if s == "":
        return 0.0

    negativo_parentesis = s.startswith("(") and s.endswith(")")
    if negativo_parentesis:
        s = s[1:-1].strip()

    s = s.replace("$", "").replace(",", "").replace(" ", "")
    s = s.replace("MXN", "").replace("USD", "")

    # Guiones aislados suelen representar vacío.
    if s in {"-", "—", "–"}:
        return 0.0

    try:
        n = float(s)
        return -n if negativo_parentesis else n
    except ValueError:
        return np.nan


def columna_a_monto(serie):
    return serie.apply(parse_amount)


def es_vacio(valor):
    return pd.isna(valor) or str(valor).strip() == ""


def extraer_numero_folio(ref_norm):
    if not ref_norm:
        return None
    m = re.search(r"(\d+)$", str(ref_norm))
    return m.group(1) if m else None


def normalizar_referencia_base(ref):
    """
    Normaliza SIN destruir prefijos documentales.

    Ejemplos:
      NCTA12846    -> NCTA12846
      NCTA-13547   -> NCTA13547
      NCTA 14314   -> NCTA14314
      NC11404      -> NC11404
      FACTURA NCTA-13547 -> NCTA13547
      8729         -> 8729

    Referencias libres se conservan en forma normalizada, pero NO se marcan
    automáticamente como folios.
    """
    if es_vacio(ref):
        return None, "VACIA", None

    # Evitar 123.0 cuando Excel leyó un folio numérico como float.
    if isinstance(ref, float) and ref.is_integer():
        s = str(int(ref))
    else:
        s = str(ref).strip()

    s = texto_norm(s)
    s = re.sub(
        r"^(?:FACTURA|FAC|FOLIO|REF|REFERENCIA)\s*[:.\-]?\s*",
        "",
        s,
    )

    # Folio con prefijo conocido.
    prefijos = "|".join(sorted(PREFIJOS_FOLIO, key=len, reverse=True))
    m = re.fullmatch(rf"({prefijos})[\s\-_/.:]*(\d+)", s)
    if m:
        pref, num = m.groups()
        return f"{pref}{num}", "FOLIO_PREFIJO", pref

    # Folio completamente numérico.
    if re.fullmatch(r"\d+", s):
        return s, "FOLIO_NUMERICO", None

    # Referencia libre / bancaria / descriptiva.
    libre = re.sub(r"\s+", " ", s).strip()
    return libre if libre else None, "OTRA_REFERENCIA", None


def extraer_referencia_de_concepto(concepto):
    """
    Recupera un folio del concepto SOLO cuando existe un candidato documental
    inequívoco. No intenta inferir folios a partir de cualquier número.
    """
    s = texto_norm(concepto)
    if not s:
        return None

    prefijos = "|".join(sorted(PREFIJOS_FOLIO, key=len, reverse=True))
    patron = re.compile(
        rf"(?<![A-Z0-9])({prefijos})[\s\-_/.:]*(\d{{2,}})(?!\d)"
    )
    encontrados = {
        f"{m.group(1)}{m.group(2)}" for m in patron.finditer(s)
    }
    if len(encontrados) == 1:
        return next(iter(encontrados))
    return None


def enriquecer_referencias(df):
    df = df.copy()
    df["referencia_original"] = df["referencia"]

    refs_norm = []
    refs_tipo = []
    refs_prefijo = []
    refs_fuente = []
    refs_recuperadas = []

    for _, row in df.iterrows():
        original = row["referencia_original"]
        norm, tipo, prefijo = normalizar_referencia_base(original)
        recuperada = False
        fuente = "original"

        if norm is None:
            rec = extraer_referencia_de_concepto(row.get("concepto", ""))
            if rec:
                norm, tipo, prefijo = normalizar_referencia_base(rec)
                fuente = "concepto"
                recuperada = True
            else:
                fuente = "vacia"

        refs_norm.append(norm)
        refs_tipo.append(tipo)
        refs_prefijo.append(prefijo)
        refs_fuente.append(fuente)
        refs_recuperadas.append(recuperada)

    df["referencia_norm"] = refs_norm
    df["referencia_tipo"] = refs_tipo
    df["referencia_prefijo"] = refs_prefijo
    df["referencia_fuente"] = refs_fuente
    df["referencia_recuperada"] = refs_recuperadas
    df["referencia_numero"] = df["referencia_norm"].apply(extraer_numero_folio)

    df["tiene_referencia"] = df["referencia_norm"].notna()
    df["es_folio"] = df["referencia_tipo"].isin(
        ["FOLIO_PREFIJO", "FOLIO_NUMERICO"]
    )
    return df


def cargar_archivo_robusto(file_bytes, file_name):
    """
    Lee Excel o CSV. CSV intenta UTF-8 antes de latin-1 para evitar mojibake.
    """
    bio = BytesIO(file_bytes)
    lower = file_name.lower()

    if lower.endswith((".xlsx", ".xlsm", ".xls")):
        return pd.read_excel(bio, header=None)

    # Incluso si la extensión es CSV, primero intentamos Excel por seguridad
    # cuando el contenido realmente lo sea.
    try:
        return pd.read_excel(bio, header=None)
    except Exception:
        pass

    errores = []
    for encoding in ("utf-8-sig", "utf-8", "latin-1"):
        try:
            bio.seek(0)
            return pd.read_csv(
                bio,
                header=None,
                encoding=encoding,
                sep=None,
                engine="python",
            )
        except Exception as e:
            errores.append(f"{encoding}: {e}")

    raise ValueError(
        "No fue posible leer el archivo como Excel ni CSV. "
        + " | ".join(errores[-2:])
    )


def to_excel_workbook(tablas):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        usados = set()
        for nombre, df in tablas.items():
            hoja = re.sub(r"[\[\]\*\?/\\:]", "_", str(nombre))[:31] or "Datos"
            base = hoja
            i = 2
            while hoja in usados:
                suf = f"_{i}"
                hoja = (base[: 31 - len(suf)] + suf)
                i += 1
            usados.add(hoja)
            (df if df is not None else pd.DataFrame()).to_excel(
                writer, sheet_name=hoja, index=False
            )
    return output.getvalue()


def buscar_saldo_inicial_en_fila(row):
    vals = list(row)
    for i, valor in enumerate(vals):
        if "saldo inicial" in texto_norm(valor).lower():
            for j in range(i + 1, len(vals)):
                n = parse_amount(vals[j])
                if not pd.isna(n):
                    return float(n)
    return np.nan


def buscar_columna_texto(row, objetivo_exacto):
    objetivo = texto_norm(objetivo_exacto)
    for i, valor in enumerate(row):
        if texto_norm(valor) == objetivo:
            return i
    return None


# ==============================================================================
# 2. MOTOR DE LECTURA Y VALIDACIÓN
# ==============================================================================

def procesar_archivo_core(file_bytes, file_name):
    raw = cargar_archivo_robusto(file_bytes, file_name)

    if raw.shape[1] < 8:
        raise ValueError(
            f"{file_name}: se esperaban al menos 8 columnas del auxiliar CONTPAQ "
            f"y se encontraron {raw.shape[1]}."
        )

    raw_str = raw.fillna("").astype(str)

    # Más flexible que NNN-NNN-NNN-NNN, pero sigue exigiendo estructura de cuenta.
    patron_cuenta = r"^\d+(?:-\d+){2,}$"
    mask_candidato_cuenta = raw_str[0].str.strip().str.match(
        patron_cuenta, na=False
    )

    mask_saldo = raw_str.apply(
        lambda r: r.str.contains("Saldo inicial", case=False, na=False).any(),
        axis=1,
    )
    is_header = mask_candidato_cuenta & mask_saldo

    candidatos_no_header = raw.index[
        mask_candidato_cuenta & ~is_header
    ].tolist()

    # Movimientos: primero reconocemos la forma; después validamos que la fecha
    # realmente pueda convertirse.
    patron_fecha = (
        r"^\d{1,2}[/\-][A-Za-zÁÉÍÓÚáéíóúÑñ]{3,4}[/\-]\d{4}$"
    )
    is_mov = raw_str[0].str.strip().str.match(patron_fecha, na=False)

    # Totales por cuenta. En este formato aparecen en columna E (índice 4).
    mask_total = raw_str[4].str.strip().eq("Total:")

    # Total general espaciado: "T o t a l:".
    col4_compacto = (
        raw_str[4]
        .str.replace(r"\s+", "", regex=True)
        .str.strip()
        .str.lower()
    )
    es_total_general = (
        col4_compacto.eq("total:")
        & ~raw_str[4].str.strip().eq("Total:")
    )

    n_headers = int(is_header.sum())
    n_totales = int(mask_total.sum())
    n_movs = int(is_mov.sum())

    problemas_estructura = []
    if n_headers == 0:
        problemas_estructura.append(
            "No se detectó ningún encabezado de cuenta con saldo inicial."
        )
    if candidatos_no_header:
        problemas_estructura.append(
            f"Hay {len(candidatos_no_header)} fila(s) que parecen código de cuenta "
            "pero no contienen 'Saldo inicial'."
        )
    if n_headers != n_totales:
        problemas_estructura.append(
            f"Encabezados de cuenta ({n_headers}) != filas 'Total:' ({n_totales})."
        )
    if n_movs == 0:
        problemas_estructura.append("No se detectaron movimientos con fecha dd/Mmm/aaaa.")

    if problemas_estructura:
        raise ValueError(
            f"{file_name}: la estructura no puede certificarse. "
            + " ".join(problemas_estructura)
        )

    df = raw.copy()
    df["meta_codigo"] = np.where(is_header, raw[0], np.nan)
    df["meta_nombre"] = np.where(is_header, raw[1], np.nan)

    saldo_ini_map = {}
    for idx in raw.index[is_header]:
        saldo_ini = buscar_saldo_inicial_en_fila(raw.loc[idx])
        if pd.isna(saldo_ini):
            raise ValueError(
                f"{file_name}: no pude leer el saldo inicial de la cuenta "
                f"{raw.loc[idx, 0]} (fila Excel {idx + 1})."
            )
        saldo_ini_map[idx] = saldo_ini

    df["meta_saldo_inicial"] = np.nan
    for idx, saldo in saldo_ini_map.items():
        df.loc[idx, "meta_saldo_inicial"] = saldo

    # El ffill solo se hace DESPUÉS de validar que todos los candidatos de cuenta
    # fueron reconocidos como headers.
    df["meta_codigo"] = df["meta_codigo"].ffill()
    df["meta_nombre"] = df["meta_nombre"].ffill()
    df["meta_saldo_inicial"] = df["meta_saldo_inicial"].ffill()

    # Verifica movimientos huérfanos.
    if df.loc[is_mov, "meta_codigo"].isna().any():
        filas = (df.index[is_mov & df["meta_codigo"].isna()] + 1).tolist()
        raise ValueError(
            f"{file_name}: hay movimientos sin cuenta asociada en filas {filas[:10]}."
        )

    # --------------------------------------------------------------------------
    # Movimientos
    # --------------------------------------------------------------------------
    movs = df[is_mov].copy()
    movs = movs.rename(
        columns={
            0: "fecha_raw",
            1: "tipo_poliza",
            2: "poliza",
            3: "concepto",
            4: "referencia",
            5: "cargos",
            6: "abonos",
            7: "saldo_acumulado",
        }
    )

    movs["fila_origen"] = movs.index + 1
    movs["archivo"] = file_name
    movs["cuenta_uid"] = (
        movs["archivo"].astype(str) + "::" + movs["meta_codigo"].astype(str)
    )

    # Fechas
    movs["fecha"] = movs["fecha_raw"].apply(parse_spanish_date)
    if movs["fecha"].isna().any():
        filas = movs.loc[movs["fecha"].isna(), "fila_origen"].tolist()
        raise ValueError(
            f"{file_name}: {len(filas)} fecha(s) de movimiento no pudieron "
            f"interpretarse. Filas: {filas[:10]}."
        )

    # Montos: no convertir texto inválido a cero.
    for c in ["cargos", "abonos", "saldo_acumulado"]:
        original = movs[c].copy()
        convertido = columna_a_monto(original)
        invalidos = convertido.isna()
        if invalidos.any():
            filas = movs.loc[invalidos, "fila_origen"].tolist()
            ejemplos = original[invalidos].astype(str).head(5).tolist()
            raise ValueError(
                f"{file_name}: valores no numéricos en '{c}' en "
                f"{len(filas)} movimiento(s). Filas {filas[:10]}; "
                f"ejemplos: {ejemplos}."
            )
        movs[c] = convertido.astype(float)

    movs["concepto_norm"] = movs["concepto"].apply(concepto_norm)
    movs = enriquecer_referencias(movs)

    # --------------------------------------------------------------------------
    # Totales por cuenta
    # --------------------------------------------------------------------------
    totales = df[mask_total].copy()
    totales["archivo"] = file_name
    totales["cuenta_uid"] = (
        totales["archivo"].astype(str)
        + "::"
        + totales["meta_codigo"].astype(str)
    )

    for col_idx, nombre in [
        (5, "total_cargos"),
        (6, "total_abonos"),
        (7, "saldo_final_aux"),
    ]:
        totales[nombre] = columna_a_monto(totales[col_idx])
        if totales[nombre].isna().any():
            filas = (totales.index[totales[nombre].isna()] + 1).tolist()
            raise ValueError(
                f"{file_name}: no pude leer '{nombre}' en Total:. "
                f"Filas {filas[:10]}."
            )

    # Para este tipo de reporte exigimos un total por cuenta.
    total_por_cuenta = totales.groupby("cuenta_uid").size()
    multiples = total_por_cuenta[total_por_cuenta != 1]
    if not multiples.empty:
        raise ValueError(
            f"{file_name}: se esperaba exactamente un 'Total:' por cuenta. "
            f"Casos: {multiples.to_dict()}."
        )

    resumen = totales[
        [
            "archivo", "cuenta_uid", "meta_codigo", "meta_nombre",
            "meta_saldo_inicial", "total_cargos", "total_abonos",
            "saldo_final_aux"
        ]
    ].copy()

    resumen = resumen.rename(
        columns={"meta_saldo_inicial": "saldo_inicial"}
    )

    # --------------------------------------------------------------------------
    # Gran total del reporte
    # --------------------------------------------------------------------------
    gran_total = None
    if es_total_general.any():
        idx_gt = es_total_general[es_total_general].index[-1]
        gran_total = parse_amount(raw.iloc[idx_gt, 7])
        if pd.isna(gran_total):
            gran_total = None

    # Comparación entre suma de movimientos y Total: de CONTPAQ.
    sum_mov = (
        movs.groupby("cuenta_uid", as_index=False)
        .agg(
            mov_cargos=("cargos", "sum"),
            mov_abonos=("abonos", "sum"),
        )
    )
    resumen = resumen.merge(sum_mov, on="cuenta_uid", how="left")
    resumen[["mov_cargos", "mov_abonos"]] = resumen[
        ["mov_cargos", "mov_abonos"]
    ].fillna(0.0)

    resumen["dif_cargos_vs_total"] = (
        resumen["total_cargos"] - resumen["mov_cargos"]
    )
    resumen["dif_abonos_vs_total"] = (
        resumen["total_abonos"] - resumen["mov_abonos"]
    )

    # Si el propio Total: de cargos/abonos no coincide con el detalle, no
    # podemos declarar una lectura "verificada".
    mal_detalle = resumen[
        (resumen["dif_cargos_vs_total"].abs() > UMBRAL_TOLERANCIA)
        | (resumen["dif_abonos_vs_total"].abs() > UMBRAL_TOLERANCIA)
    ]
    if not mal_detalle.empty:
        raise ValueError(
            f"{file_name}: los cargos/abonos del detalle no amarran con el "
            "'Total:' de CONTPAQ en una o más cuentas. No se certifica la lectura."
        )

    diag = {
        "archivo": file_name,
        "n_headers": n_headers,
        "n_totales": n_totales,
        "n_movs": n_movs,
        "gran_total": float(gran_total) if gran_total is not None else None,
        "suma_saldos_cuenta": float(resumen["saldo_final_aux"].sum()),
        "n_candidatos_no_header": len(candidatos_no_header),
    }

    if gran_total is not None:
        diag["amarre_gran_total"] = (
            abs(diag["suma_saldos_cuenta"] - gran_total)
            <= max(UMBRAL_TOLERANCIA, abs(gran_total) * 1e-6)
        )
    else:
        diag["amarre_gran_total"] = None

    return movs.reset_index(drop=True), resumen.reset_index(drop=True), diag


@st.cache_data(show_spinner=False)
def procesar_archivo_engine(file_bytes, file_name):
    return procesar_archivo_core(file_bytes, file_name)


# ==============================================================================
# 3. NATURALEZA CONTABLE Y CONCILIACIÓN
# ==============================================================================

def detectar_naturaleza(resumen, movs):
    """
    Detecta naturaleza por cuenta comparando las dos ecuaciones posibles
    contra el saldo final que reporta CONTPAQ.

    Deudora:
      final = inicial + cargos - abonos

    Acreedora:
      final = inicial - cargos + abonos

    También usa el comportamiento del saldo acumulado como evidencia secundaria.
    """
    r = resumen.copy()

    r["esperado_deudora"] = (
        r["saldo_inicial"] + r["total_cargos"] - r["total_abonos"]
    )
    r["esperado_acreedora"] = (
        r["saldo_inicial"] - r["total_cargos"] + r["total_abonos"]
    )
    r["error_deudora"] = r["saldo_final_aux"] - r["esperado_deudora"]
    r["error_acreedora"] = r["saldo_final_aux"] - r["esperado_acreedora"]

    naturalezas = []
    confianzas = []

    for _, row in r.iterrows():
        ed = abs(row["error_deudora"])
        ea = abs(row["error_acreedora"])

        deud_cuadra = ed <= UMBRAL_TOLERANCIA
        acre_cuadra = ea <= UMBRAL_TOLERANCIA

        if deud_cuadra and not acre_cuadra:
            naturaleza = "DEUDORA"
            confianza = "ALTA"
        elif acre_cuadra and not deud_cuadra:
            naturaleza = "ACREEDORA"
            confianza = "ALTA"
        elif deud_cuadra and acre_cuadra:
            # Puede ocurrir si cargos == abonos. Revisamos el saldo acumulado.
            mm = movs[movs["cuenta_uid"] == row["cuenta_uid"]].copy()
            if len(mm):
                mm = mm.sort_values("fila_origen")
                saldo_previo = mm["saldo_acumulado"].shift(1)
                saldo_previo.iloc[0] = row["saldo_inicial"]
                delta_obs = mm["saldo_acumulado"] - saldo_previo

                err_mov_deud = (
                    delta_obs - (mm["cargos"] - mm["abonos"])
                ).abs().sum()
                err_mov_acre = (
                    delta_obs - (mm["abonos"] - mm["cargos"])
                ).abs().sum()

                if err_mov_deud + 0.01 < err_mov_acre:
                    naturaleza = "DEUDORA"
                    confianza = "ALTA"
                elif err_mov_acre + 0.01 < err_mov_deud:
                    naturaleza = "ACREEDORA"
                    confianza = "ALTA"
                else:
                    naturaleza = "INDETERMINADA"
                    confianza = "BAJA"
            else:
                naturaleza = "INDETERMINADA"
                confianza = "BAJA"
        else:
            # Ninguna ecuación cuadra. Elegimos solo si una es claramente mejor;
            # de cualquier forma la confianza será baja y quedará hallazgo.
            if ed < ea * 0.25:
                naturaleza = "DEUDORA"
                confianza = "BAJA"
            elif ea < ed * 0.25:
                naturaleza = "ACREEDORA"
                confianza = "BAJA"
            else:
                naturaleza = "INDETERMINADA"
                confianza = "BAJA"

        naturalezas.append(naturaleza)
        confianzas.append(confianza)

    r["naturaleza"] = naturalezas
    r["naturaleza_confianza"] = confianzas
    return r


def aplicar_naturaleza_a_movimientos(movs, resumen_naturaleza):
    m = movs.copy()
    mapa_nat = resumen_naturaleza.set_index("cuenta_uid")["naturaleza"]
    m["naturaleza"] = m["cuenta_uid"].map(mapa_nat)

    m["efecto_natural"] = np.select(
        [
            m["naturaleza"].eq("DEUDORA"),
            m["naturaleza"].eq("ACREEDORA"),
        ],
        [
            m["cargos"] - m["abonos"],
            m["abonos"] - m["cargos"],
        ],
        default=np.nan,
    )

    m["importe_abs"] = m["efecto_natural"].abs().round(2)
    return m


def marcar_duplicados_exactos(movs):
    m = movs.copy()
    subset = [
        "cuenta_uid", "fecha", "tipo_poliza", "poliza", "concepto_norm",
        "referencia_norm", "cargos", "abonos"
    ]
    m["posible_duplicado_exacto"] = m.duplicated(
        subset=subset, keep=False
    )
    return m


def analizar_saldos(movs, resumen_naturaleza):
    """
    Conciliación por naturaleza:

      saldo_final =
          saldo_inicial
        + efecto natural de movimientos CON referencia
        + efecto natural de movimientos SIN referencia
        + descuadre_origen

    Los hallazgos son independientes; el estado es solo una prioridad visual.
    """
    m = movs.copy()
    r = resumen_naturaleza.copy()

    con_ref = (
        m[m["tiene_referencia"]]
        .groupby("cuenta_uid")["efecto_natural"]
        .sum(min_count=1)
    )
    sin_ref = (
        m[~m["tiene_referencia"]]
        .groupby("cuenta_uid")["efecto_natural"]
        .sum(min_count=1)
    )

    sin_ref_stats = (
        m[~m["tiene_referencia"]]
        .groupby("cuenta_uid")
        .agg(
            n_sin_referencia=("cuenta_uid", "size"),
            cargos_sin_referencia=("cargos", "sum"),
            abonos_sin_referencia=("abonos", "sum"),
        )
    )

    rec_stats = (
        m[m["referencia_recuperada"]]
        .groupby("cuenta_uid")
        .size()
        .rename("n_refs_recuperadas")
    )

    otras_ref_stats = (
        m[m["referencia_tipo"].eq("OTRA_REFERENCIA")]
        .groupby("cuenta_uid")
        .size()
        .rename("n_referencias_libres")
    )

    neg_stats = (
        m[(m["cargos"] < 0) | (m["abonos"] < 0)]
        .groupby("cuenta_uid")
        .size()
        .rename("n_montos_negativos")
    )

    dup_stats = (
        m[m["posible_duplicado_exacto"]]
        .groupby("cuenta_uid")
        .size()
        .rename("n_filas_posible_duplicado")
    )

    r["movs_con_referencia"] = r["cuenta_uid"].map(con_ref).fillna(0.0)
    r["movs_sin_referencia"] = r["cuenta_uid"].map(sin_ref).fillna(0.0)

    r = r.merge(
        sin_ref_stats,
        left_on="cuenta_uid",
        right_index=True,
        how="left",
    )
    for c in [
        "n_sin_referencia", "cargos_sin_referencia",
        "abonos_sin_referencia"
    ]:
        r[c] = r[c].fillna(0)

    r["importe_bruto_sin_referencia"] = (
        r["cargos_sin_referencia"].abs()
        + r["abonos_sin_referencia"].abs()
    )

    r["n_refs_recuperadas"] = r["cuenta_uid"].map(rec_stats).fillna(0).astype(int)
    r["n_referencias_libres"] = (
        r["cuenta_uid"].map(otras_ref_stats).fillna(0).astype(int)
    )
    r["n_montos_negativos"] = (
        r["cuenta_uid"].map(neg_stats).fillna(0).astype(int)
    )
    r["n_filas_posible_duplicado"] = (
        r["cuenta_uid"].map(dup_stats).fillna(0).astype(int)
    )

    r["saldo_esperado_motor"] = (
        r["saldo_inicial"]
        + r["movs_con_referencia"]
        + r["movs_sin_referencia"]
    )
    r["descuadre_origen"] = (
        r["saldo_final_aux"] - r["saldo_esperado_motor"]
    )

    r["cuadra"] = (
        r["descuadre_origen"].abs() <= UMBRAL_TOLERANCIA
    )
    r["tiene_arrastre"] = (
        r["saldo_inicial"].abs() > UMBRAL_TOLERANCIA
    )
    r["tiene_sin_referencia"] = r["n_sin_referencia"] > 0
    r["tiene_montos_negativos"] = r["n_montos_negativos"] > 0

    def estado(row):
        if row["naturaleza"] == "INDETERMINADA":
            return "⚫ Naturaleza indeterminada"
        if abs(row["descuadre_origen"]) > UMBRAL_TOLERANCIA:
            return "🟠 Total CONTPAQ ≠ Detalle"
        if row["n_sin_referencia"] > 0:
            return "🔴 Movimientos sin referencia"
        if row["n_montos_negativos"] > 0:
            return "🟣 Montos negativos / reversos"
        return "🟢 OK"

    r["estado"] = r.apply(estado, axis=1)
    return r


# ==============================================================================
# 4. FOLIOS, REFERENCIAS Y CRUCES
# ==============================================================================

def analizar_folios(movs, fecha_corte):
    """
    Analiza solo referencias que parecen folio documental.
    La antigüedad es desde la PRIMERA fecha observada, NO fecha de vencimiento.
    """
    mv = movs[
        movs["es_folio"]
        & movs["efecto_natural"].notna()
    ].copy()

    if mv.empty:
        return pd.DataFrame(
            columns=[
                "archivo", "meta_codigo", "meta_nombre", "naturaleza",
                "referencia_norm", "primera_fecha", "ultima_fecha", "n_movs",
                "cargos", "abonos", "saldo_natural", "dias",
                "antiguedad_observada", "tipo_saldo",
                "multiples_movimientos", "posible_duplicado_exacto"
            ]
        )

    g = (
        mv.groupby(
            [
                "archivo", "cuenta_uid", "meta_codigo", "meta_nombre",
                "naturaleza", "referencia_norm"
            ],
            as_index=False,
        )
        .agg(
            primera_fecha=("fecha", "min"),
            ultima_fecha=("fecha", "max"),
            n_movs=("efecto_natural", "size"),
            cargos=("cargos", "sum"),
            abonos=("abonos", "sum"),
            saldo_natural=("efecto_natural", "sum"),
            posible_duplicado_exacto=(
                "posible_duplicado_exacto", "max"
            ),
        )
    )

    vivos = g[g["saldo_natural"].abs() > UMBRAL_FOLIO].copy()

    corte = pd.Timestamp(fecha_corte)
    vivos["dias"] = (corte - vivos["primera_fecha"]).dt.days

    def bucket(d):
        if pd.isna(d):
            return "sin fecha"
        if d < 0:
            return "fecha posterior al corte"
        if d <= 30:
            return "0-30"
        if d <= 60:
            return "31-60"
        if d <= 90:
            return "61-90"
        return "90+"

    vivos["antiguedad_observada"] = vivos["dias"].apply(bucket)

    def tipo_saldo(row):
        s = row["saldo_natural"]
        nat = row["naturaleza"]
        if nat == "DEUDORA":
            if s > 0:
                return "🔵 Saldo deudor pendiente"
            return "🔴 Saldo contrario a naturaleza (acreedor)"
        if nat == "ACREEDORA":
            if s > 0:
                return "🟣 Saldo acreedor pendiente / por aplicar"
            return "🔴 Saldo contrario a naturaleza (deudor)"
        return "⚫ Naturaleza indeterminada"

    vivos["tipo_saldo"] = vivos.apply(tipo_saldo, axis=1)
    vivos["multiples_movimientos"] = vivos["n_movs"] > 2

    return vivos.sort_values(
        ["dias", "saldo_natural"], ascending=[False, False]
    )


def detectar_cruces_por_referencia(movs):
    """
    Misma referencia en MÁS DE UNA cuenta y efectos naturales opuestos.
    """
    mv = movs[
        movs["es_folio"]
        & movs["efecto_natural"].notna()
    ].copy()

    if mv.empty:
        return pd.DataFrame()

    por_cuenta = (
        mv.groupby(
            [
                "referencia_norm", "meta_codigo", "meta_nombre",
                "naturaleza"
            ],
            as_index=False,
        )
        .agg(
            cargos=("cargos", "sum"),
            abonos=("abonos", "sum"),
            efecto_natural=("efecto_natural", "sum"),
            n_movs=("efecto_natural", "size"),
        )
    )

    nivel_ref = (
        por_cuenta.groupby("referencia_norm")
        .agg(
            num_cuentas=("meta_codigo", "nunique"),
            hay_positivo=("efecto_natural", lambda x: (x > UMBRAL_FOLIO).any()),
            hay_negativo=("efecto_natural", lambda x: (x < -UMBRAL_FOLIO).any()),
            neto_global=("efecto_natural", "sum"),
        )
        .reset_index()
    )

    refs = nivel_ref[
        (nivel_ref["num_cuentas"] > 1)
        & nivel_ref["hay_positivo"]
        & nivel_ref["hay_negativo"]
    ].copy()

    if refs.empty:
        return pd.DataFrame()

    detalle = por_cuenta[
        por_cuenta["referencia_norm"].isin(refs["referencia_norm"])
    ].merge(
        refs[["referencia_norm", "num_cuentas", "neto_global"]],
        on="referencia_norm",
        how="left",
    )

    detalle["amarre_aprox"] = (
        detalle["neto_global"].abs() <= UMBRAL_TOLERANCIA
    )
    return detalle.sort_values(
        ["referencia_norm", "efecto_natural"], ascending=[True, False]
    )


def detectar_coincidencias_por_evidencia(movs):
    """
    Cruces fuertes aunque el folio sea diferente:
      misma fecha + mismo concepto normalizado + mismo importe absoluto
      + cuentas distintas + efectos naturales opuestos.

    Se trabaja por grupos, no por similitud difusa de nombres.
    """
    mv = movs[
        movs["efecto_natural"].notna()
        & (movs["importe_abs"] > UMBRAL_FOLIO)
        & movs["concepto_norm"].ne("")
    ].copy()

    if mv.empty:
        return pd.DataFrame()

    # Excluir movimientos que tienen cargo y abono simultáneamente y netean 0.
    mv = mv[mv["efecto_natural"].abs() > UMBRAL_FOLIO].copy()

    claves = ["fecha", "concepto_norm", "importe_abs"]

    grupos = (
        mv.groupby(claves)
        .agg(
            num_cuentas=("meta_codigo", "nunique"),
            hay_positivo=("efecto_natural", lambda x: (x > 0).any()),
            hay_negativo=("efecto_natural", lambda x: (x < 0).any()),
            n_movs_grupo=("efecto_natural", "size"),
            neto_grupo=("efecto_natural", "sum"),
        )
        .reset_index()
    )

    validos = grupos[
        (grupos["num_cuentas"] > 1)
        & grupos["hay_positivo"]
        & grupos["hay_negativo"]
    ].copy()

    if validos.empty:
        return pd.DataFrame()

    validos["evidencia_id"] = np.arange(1, len(validos) + 1)

    det = mv.merge(validos, on=claves, how="inner")
    cols = [
        "evidencia_id", "fecha", "concepto", "concepto_norm", "importe_abs",
        "archivo", "meta_codigo", "meta_nombre", "naturaleza",
        "referencia_original", "referencia_norm", "referencia_fuente",
        "cargos", "abonos", "efecto_natural",
        "num_cuentas", "n_movs_grupo", "neto_grupo"
    ]
    return det[cols].sort_values(
        ["evidencia_id", "efecto_natural"], ascending=[True, False]
    )


def tabla_referencias(movs):
    cols = [
        "archivo", "fila_origen", "fecha", "meta_codigo", "meta_nombre",
        "concepto", "referencia_original", "referencia_norm",
        "referencia_tipo", "referencia_fuente", "referencia_recuperada",
        "cargos", "abonos", "naturaleza", "efecto_natural"
    ]
    return movs[cols].copy()


# ==============================================================================
# 5. UI
# ==============================================================================

def main():
    st.set_page_config(
        page_title="Auditoría Master CONTPAQ",
        layout="wide",
        page_icon="🛡️",
    )

    st.title("🛡️ Auditoría Master de Saldos (CONTPAQ)")
    st.caption(f"Motor v{APP_VERSION}")

    st.markdown(
        """
        Esta versión:
        - valida que el auxiliar haya sido leído íntegramente;
        - detecta automáticamente **naturaleza deudora o acreedora**;
        - conserva el folio original y normaliza sin destruir prefijos;
        - puede recuperar un folio desde **Concepto** cuando Referencia está vacía;
        - separa referencias libres de folios documentales;
        - detecta movimientos sin referencia por **cantidad, cargos, abonos y bruto**, no solo por neto;
        - identifica montos negativos/reversos y posibles duplicados exactos;
        - con varios archivos, busca cruces por referencia y por evidencia contable.
        """
    )

    uploaded_files = st.file_uploader(
        "📂 Sube uno o varios auxiliares CONTPAQ (Excel o CSV)",
        type=["xlsx", "xls", "xlsm", "csv"],
        accept_multiple_files=True,
    )

    if not uploaded_files:
        st.info("Esperando archivo(s)...")
        return

    movs_lista = []
    resumen_lista = []
    diags = []
    errores = []

    with st.spinner("Procesando y validando auxiliares..."):
        for uf in uploaded_files:
            try:
                movs_i, resumen_i, diag_i = procesar_archivo_engine(
                    uf.getvalue(), uf.name
                )
                movs_lista.append(movs_i)
                resumen_lista.append(resumen_i)
                diags.append(diag_i)
            except Exception as e:
                errores.append(f"**{uf.name}:** {e}")

    if errores:
        st.error(
            "⚠️ **No se certifica la lectura. Corrige o revisa estos archivos antes "
            "de usar resultados:**\n\n"
            + "\n\n".join(f"- {x}" for x in errores)
        )
        st.stop()

    movs = pd.concat(movs_lista, ignore_index=True)
    resumen = pd.concat(resumen_lista, ignore_index=True)

    # Evitar que cargar dos veces la misma cuenta pase desapercibido.
    repetidas = (
        resumen.groupby("meta_codigo")["archivo"]
        .nunique()
        .loc[lambda s: s > 1]
    )
    if not repetidas.empty:
        st.warning(
            "⚠️ Hay códigos de cuenta presentes en más de un archivo. "
            "No necesariamente es un error, pero revisa que no hayas cargado "
            "dos periodos o copias de la misma cuenta: "
            + ", ".join(repetidas.index.astype(str))
        )

    resumen_nat = detectar_naturaleza(resumen, movs)
    movs = aplicar_naturaleza_a_movimientos(movs, resumen_nat)
    movs = marcar_duplicados_exactos(movs)
    df_audit = analizar_saldos(movs, resumen_nat)

    # --------------------------------------------------------------------------
    # Validación visible de lectura
    # --------------------------------------------------------------------------
    st.divider()
    st.subheader("✅ Validación de lectura")

    diag_df = pd.DataFrame(diags)
    n_archivos = len(diag_df)
    n_cuentas = len(df_audit)
    n_movs = len(movs)

    amarres_false = diag_df["amarre_gran_total"].eq(False).sum()
    if amarres_false:
        st.warning(
            f"{amarres_false} archivo(s) no amarran la suma de saldos por cuenta "
            "contra el Total general detectado. Revisa si el reporte contiene "
            "agrupaciones adicionales."
        )
    else:
        st.success(
            f"Lectura estructural validada: **{n_archivos} archivo(s)** · "
            f"**{n_cuentas} cuenta(s)** · **{n_movs:,} movimientos**. "
            "Los cargos y abonos del detalle amarran con cada 'Total:' de CONTPAQ."
        )

    for _, d in diag_df.iterrows():
        gt_txt = (
            f"${d['gran_total']:,.2f}"
            if pd.notna(d.get("gran_total"))
            else "no detectado"
        )
        amarre = d.get("amarre_gran_total")
        if pd.isna(amarre):
            estado_amarre = "ℹ️"
        elif bool(amarre):
            estado_amarre = "✅"
        else:
            estado_amarre = "⚠️"
        st.caption(
            f"{estado_amarre} {d['archivo']}: "
            f"{int(d['n_headers'])} cuenta(s), {int(d['n_movs']):,} movimientos, "
            f"Total general {gt_txt}."
        )

    # --------------------------------------------------------------------------
    # KPIs
    # --------------------------------------------------------------------------
    saldo_total = df_audit["saldo_final_aux"].sum()
    bruto_sin_ref = df_audit["importe_bruto_sin_referencia"].sum()
    descuadre_abs = df_audit["descuadre_origen"].abs().sum()
    n_sin_ref = int(df_audit["n_sin_referencia"].sum())
    n_revisar = int((df_audit["estado"] != "🟢 OK").sum())

    df_cruces_ref = detectar_cruces_por_referencia(movs)
    df_evidencia = detectar_coincidencias_por_evidencia(movs)

    n_refs_cruce = (
        int(df_cruces_ref["referencia_norm"].nunique())
        if not df_cruces_ref.empty else 0
    )
    n_evidencias = (
        int(df_evidencia["evidencia_id"].nunique())
        if not df_evidencia.empty else 0
    )

    k1, k2, k3, k4, k5, k6 = st.columns(6)
    k1.metric("Saldo total reportado", f"${saldo_total:,.2f}")
    k2.metric(
        "Movs sin referencia",
        f"{n_sin_ref:,}",
        help=f"Importe bruto involucrado: ${bruto_sin_ref:,.2f}",
    )
    k3.metric(
        "Descuadre absoluto",
        f"${descuadre_abs:,.2f}",
        help="Suma de valores absolutos por cuenta; evita compensar + y -.",
    )
    k4.metric("Cruces por folio", n_refs_cruce)
    k5.metric("Cruces por evidencia", n_evidencias)
    k6.metric("Cuentas a revisar", n_revisar)

    # Fecha de corte
    fmax = movs["fecha"].max()
    corte_default = (
        fmax.date() if pd.notna(fmax) else pd.Timestamp.now().date()
    )
    corte = st.date_input(
        "📅 Fecha de corte para antigüedad observada",
        value=corte_default,
        help=(
            "La antigüedad se calcula desde la primera fecha observada del folio. "
            "No equivale a días vencidos si no existe fecha de vencimiento."
        ),
    )
    folios = analizar_folios(movs, corte)

    # --------------------------------------------------------------------------
    # Pestañas
    # --------------------------------------------------------------------------
    tabs = st.tabs(
        [
            "🔎 Hallazgos",
            "🚦 Semáforo",
            "📑 Folios",
            "🔀 Cruces / Conciliación",
            "🏷️ Referencias",
            "📉 Gráficos",
            "🧪 Diagnóstico",
        ]
    )

    # --------------------------------------------------------------------------
    # Hallazgos
    # --------------------------------------------------------------------------
    with tabs[0]:
        st.subheader("🔎 Hallazgos priorizados")
        st.caption(
            "Los hallazgos son independientes. Una cuenta puede tener más de uno."
        )

        sin_ref_movs = movs[~movs["tiene_referencia"]].copy()
        refs_rec = movs[movs["referencia_recuperada"]].copy()
        negativos = movs[(movs["cargos"] < 0) | (movs["abonos"] < 0)].copy()
        duplicados = movs[movs["posible_duplicado_exacto"]].copy()
        descuadres = df_audit[
            df_audit["descuadre_origen"].abs() > UMBRAL_TOLERANCIA
        ].copy()
        contrarios = folios[
            folios["tipo_saldo"].str.contains("contrario", case=False, na=False)
        ].copy()
        viejos = folios[
            folios["antiguedad_observada"].eq("90+")
        ].copy()

        h1, h2, h3, h4, h5, h6 = st.columns(6)
        h1.metric("Cuentas descuadre", len(descuadres))
        h2.metric("Movs sin ref", len(sin_ref_movs))
        h3.metric("Refs recuperadas", len(refs_rec))
        h4.metric("Montos negativos", len(negativos))
        h5.metric("Posibles duplicados", len(duplicados))
        h6.metric("Folios 90+ observados", len(viejos))

        if len(descuadres):
            st.markdown("#### 🟠 Descuadre contra el saldo final de CONTPAQ")
            st.dataframe(
                descuadres[
                    [
                        "archivo", "meta_codigo", "meta_nombre", "naturaleza",
                        "saldo_final_aux", "saldo_esperado_motor",
                        "descuadre_origen"
                    ]
                ],
                use_container_width=True,
                hide_index=True,
            )

        if len(sin_ref_movs):
            st.markdown("#### 🔴 Movimientos realmente sin referencia")
            st.caption(
                "No había referencia en la columna y tampoco fue posible recuperar "
                "un folio documental inequívoco desde Concepto."
            )
            st.dataframe(
                sin_ref_movs[
                    [
                        "archivo", "fila_origen", "fecha", "meta_codigo",
                        "concepto", "cargos", "abonos", "efecto_natural"
                    ]
                ],
                use_container_width=True,
                hide_index=True,
            )

        if len(refs_rec):
            st.markdown("#### 🟡 Referencias recuperadas desde Concepto")
            st.caption(
                "No se consideran 'sin referencia', pero se muestran para trazabilidad."
            )
            st.dataframe(
                refs_rec[
                    [
                        "archivo", "fila_origen", "fecha", "meta_codigo",
                        "concepto", "referencia_original", "referencia_norm",
                        "cargos", "abonos"
                    ]
                ],
                use_container_width=True,
                hide_index=True,
            )

        if len(negativos):
            st.markdown("#### 🟣 Montos negativos / reversos")
            st.caption(
                "Se señalan como movimiento especial; no se reinterpretan como abono."
            )
            st.dataframe(
                negativos[
                    [
                        "archivo", "fila_origen", "fecha", "meta_codigo",
                        "concepto", "referencia_norm", "cargos", "abonos",
                        "efecto_natural"
                    ]
                ],
                use_container_width=True,
                hide_index=True,
            )

        if len(duplicados):
            st.markdown("#### 🔁 Posibles duplicados exactos")
            st.caption(
                "Misma cuenta, fecha, tipo, póliza, concepto, referencia y monto. "
                "Es un indicador para revisión, no una conclusión automática."
            )
            st.dataframe(
                duplicados[
                    [
                        "archivo", "fila_origen", "fecha", "meta_codigo",
                        "tipo_poliza", "poliza", "concepto",
                        "referencia_norm", "cargos", "abonos"
                    ]
                ],
                use_container_width=True,
                hide_index=True,
            )

        if len(contrarios):
            st.markdown("#### ⚠️ Folios con saldo contrario a la naturaleza")
            st.dataframe(
                contrarios[
                    [
                        "archivo", "meta_codigo", "meta_nombre", "naturaleza",
                        "referencia_norm", "primera_fecha", "dias",
                        "cargos", "abonos", "saldo_natural", "tipo_saldo"
                    ]
                ],
                use_container_width=True,
                hide_index=True,
            )

        if not any(
            [
                len(descuadres), len(sin_ref_movs), len(refs_rec),
                len(negativos), len(duplicados), len(contrarios)
            ]
        ):
            st.success("Sin hallazgos relevantes con los criterios actuales.")

    # --------------------------------------------------------------------------
    # Semáforo
    # --------------------------------------------------------------------------
    with tabs[1]:
        st.subheader("🚦 Conciliación por cuenta")
        solo_problemas = st.toggle(
            "Ver solo cuentas con hallazgos",
            value=False,
            key="solo_problemas",
        )
        show = (
            df_audit[df_audit["estado"] != "🟢 OK"]
            if solo_problemas else df_audit
        )

        cols = [
            "archivo", "meta_codigo", "meta_nombre", "naturaleza",
            "naturaleza_confianza", "estado", "saldo_inicial",
            "total_cargos", "total_abonos", "saldo_final_aux",
            "movs_con_referencia", "movs_sin_referencia",
            "n_sin_referencia", "importe_bruto_sin_referencia",
            "n_refs_recuperadas", "n_referencias_libres",
            "n_montos_negativos", "descuadre_origen"
        ]
        st.dataframe(
            show[cols],
            use_container_width=True,
            hide_index=True,
            column_config={
                "saldo_inicial": st.column_config.NumberColumn(
                    "Saldo inicial", format="$%.2f"
                ),
                "total_cargos": st.column_config.NumberColumn(
                    "Cargos", format="$%.2f"
                ),
                "total_abonos": st.column_config.NumberColumn(
                    "Abonos", format="$%.2f"
                ),
                "saldo_final_aux": st.column_config.NumberColumn(
                    "Saldo final CONTPAQ", format="$%.2f"
                ),
                "movs_con_referencia": st.column_config.NumberColumn(
                    "Efecto con referencia", format="$%.2f"
                ),
                "movs_sin_referencia": st.column_config.NumberColumn(
                    "Efecto sin referencia", format="$%.2f"
                ),
                "importe_bruto_sin_referencia": st.column_config.NumberColumn(
                    "Bruto sin referencia", format="$%.2f"
                ),
                "descuadre_origen": st.column_config.NumberColumn(
                    "Descuadre", format="$%.2f"
                ),
            },
        )

    # --------------------------------------------------------------------------
    # Folios
    # --------------------------------------------------------------------------
    with tabs[2]:
        st.subheader("📑 Folios documentales abiertos")
        st.caption(
            "Solo incluye referencias con forma documental reconocible. "
            "La antigüedad es observada desde la primera fecha del folio, "
            "no fecha contractual de vencimiento."
        )

        orden = ["0-30", "31-60", "61-90", "90+"]
        positivos = folios[folios["saldo_natural"] > 0].copy()
        aging = (
            positivos.groupby("antiguedad_observada")["saldo_natural"]
            .agg(num_folios="count", saldo="sum")
            .reindex(orden)
            .fillna(0)
            .reset_index()
        )
        st.dataframe(aging, use_container_width=True, hide_index=True)

        if not folios.empty:
            nat_sel = st.multiselect(
                "Naturaleza",
                sorted(folios["naturaleza"].dropna().unique()),
                default=sorted(folios["naturaleza"].dropna().unique()),
            )
            edades_sel = st.multiselect(
                "Antigüedad observada",
                orden,
                default=orden,
            )
            fv = folios[
                folios["naturaleza"].isin(nat_sel)
                & folios["antiguedad_observada"].isin(edades_sel)
            ]
        else:
            fv = folios

        st.dataframe(
            fv,
            use_container_width=True,
            hide_index=True,
        )

    # --------------------------------------------------------------------------
    # Cruces / conciliación
    # --------------------------------------------------------------------------
    with tabs[3]:
        st.subheader("🔀 Cruces y conciliación entre cuentas")

        st.markdown("#### A. Cruces por el mismo folio")
        if df_cruces_ref.empty:
            st.info(
                "No se encontraron folios idénticos con efectos opuestos "
                "entre cuentas cargadas."
            )
        else:
            st.dataframe(
                df_cruces_ref,
                use_container_width=True,
                hide_index=True,
            )

        st.markdown("#### B. Coincidencias fuertes aunque el folio sea diferente")
        st.caption(
            "Misma fecha + mismo concepto + mismo importe absoluto + "
            "efecto natural opuesto entre cuentas. Es evidencia para conciliar; "
            "no se basa en similitud difusa de nombres."
        )
        if df_evidencia.empty:
            st.info(
                "No se encontraron coincidencias fuertes entre las cuentas cargadas."
            )
        else:
            st.success(
                f"Se encontraron {df_evidencia['evidencia_id'].nunique():,} "
                "grupo(s) de evidencia."
            )
            st.dataframe(
                df_evidencia,
                use_container_width=True,
                hide_index=True,
            )

    # --------------------------------------------------------------------------
    # Referencias
    # --------------------------------------------------------------------------
    with tabs[4]:
        st.subheader("🏷️ Auditoría de referencias")
        refs = tabla_referencias(movs)

        tipos = (
            refs["referencia_tipo"].fillna("VACIA").value_counts()
            .rename_axis("tipo")
            .reset_index(name="movimientos")
        )
        st.dataframe(tipos, use_container_width=True, hide_index=True)

        filtro_tipo = st.multiselect(
            "Tipo de referencia",
            sorted(refs["referencia_tipo"].dropna().unique()),
            default=sorted(refs["referencia_tipo"].dropna().unique()),
        )
        refs_show = refs[refs["referencia_tipo"].isin(filtro_tipo)]
        st.dataframe(
            refs_show,
            use_container_width=True,
            hide_index=True,
        )

    # --------------------------------------------------------------------------
    # Gráficos
    # --------------------------------------------------------------------------
    with tabs[5]:
        st.subheader("📉 Composición del saldo por naturaleza")

        saldo_ini = df_audit["saldo_inicial"].sum()
        con_ref = df_audit["movs_con_referencia"].sum()
        sin_ref = df_audit["movs_sin_referencia"].sum()
        desc = df_audit["descuadre_origen"].sum()

        fig = go.Figure(
            data=[
                go.Bar(
                    name="Saldo inicial",
                    x=["Saldo total"],
                    y=[saldo_ini],
                ),
                go.Bar(
                    name="Efecto con referencia",
                    x=["Saldo total"],
                    y=[con_ref],
                ),
                go.Bar(
                    name="Efecto sin referencia",
                    x=["Saldo total"],
                    y=[sin_ref],
                ),
                go.Bar(
                    name="Descuadre",
                    x=["Saldo total"],
                    y=[desc],
                ),
            ]
        )
        fig.update_layout(
            barmode="relative",
            title="Composición del saldo reportado",
            yaxis_title="Monto",
        )
        st.plotly_chart(fig, use_container_width=True)

    # --------------------------------------------------------------------------
    # Diagnóstico
    # --------------------------------------------------------------------------
    with tabs[6]:
        st.subheader("🧪 Diagnóstico técnico")

        st.markdown("#### Archivos")
        st.dataframe(diag_df, use_container_width=True, hide_index=True)

        st.markdown("#### Detección de naturaleza")
        st.dataframe(
            df_audit[
                [
                    "archivo", "meta_codigo", "meta_nombre", "naturaleza",
                    "naturaleza_confianza", "esperado_deudora",
                    "error_deudora", "esperado_acreedora",
                    "error_acreedora", "saldo_final_aux"
                ]
            ],
            use_container_width=True,
            hide_index=True,
        )

        st.markdown("#### Definiciones importantes")
        st.info(
            "• 'Sin referencia' significa columna Referencia vacía y sin folio "
            "recuperable del Concepto.\n\n"
            "• 'Referencia libre' significa que sí existe texto en Referencia, "
            "pero no tiene forma de folio documental.\n\n"
            "• 'Antigüedad observada' no equivale a vencimiento.\n\n"
            "• 'Posible duplicado exacto' es un indicador, no una eliminación automática."
        )

    # --------------------------------------------------------------------------
    # Exportación completa
    # --------------------------------------------------------------------------
    st.divider()
    st.subheader("⬇️ Exportación")

    export_tables = {
        "Semaforo": df_audit,
        "Folios": folios,
        "Movimientos": movs,
        "Cruces_folio": df_cruces_ref,
        "Cruces_evidencia": df_evidencia,
        "Diagnostico": diag_df,
    }
    st.download_button(
        "⬇️ Descargar auditoría completa (Excel)",
        data=to_excel_workbook(export_tables),
        file_name="auditoria_master_contpaq.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
    )


if __name__ == "__main__":
    main()
