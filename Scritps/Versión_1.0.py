import json
import calendar
import unicodedata
import base64
from pathlib import Path
import pandas as pd
import numpy as np
from datetime import date, timedelta
EXCEL_PATH  = r"C:/Users/adelarosa/Documents/Reportes/Dashboards/DashboardVentasDiarias_SkinLove/2026/10_Octubre_2026/07-10-2026/Dataset.xlsx"
OUTPUT_PATH = r"C:/Users/adelarosa/Documents/Reportes/Dashboards/DashboardVentasDiarias_SkinLove/2026/10_Octubre_2026/07-10-2026/index.html"
BOL_EXCLUIR = ["BOLEUCH", "BOLEUGDE", "BOLEUMIN"]
FECHA_BASE  = date(2026, 10, 5)
ES_CIERRE_MES = False
CLAVE_CONSOLIDADA = [2801]
LOGO_PATH = r"C:\Users\adelarosa\Documents\Reportes\Dashboards\DashboardVentasDiarias_SkinLove\Logos\logo.png"
MESES_ES = ["enero","febrero","marzo","abril","mayo","junio",
            "julio","agosto","septiembre","octubre","noviembre","diciembre"]
DIAS_ES  = ["lunes","martes","miércoles","jueves","viernes","sábado","domingo"]
MES_A_NUM = {
    "ene":1,"feb":2,"mar":3,"abr":4,"may":5,"jun":6,
    "jul":7,"ago":8,"sep":9,"oct":10,"nov":11,"dic":12,
    "enero":1,"febrero":2,"marzo":3,"abril":4,"mayo":5,"junio":6,
    "julio":7,"agosto":8,"septiembre":9,"octubre":10,"noviembre":11,"diciembre":12
}
def _margen_seguro(utilidad, ventas):
    m = (utilidad / ventas).replace([np.inf, -np.inf], 0).fillna(0)
    return m.round(4)
def _mapear_mes(serie_mes, origen: str):
    serie_norm = serie_mes.str.strip().str.lower()
    mes_num = serie_norm.map(MES_A_NUM)
    no_reconocidos = serie_norm[mes_num.isna()]
    if len(no_reconocidos):
        valores = sorted(no_reconocidos.unique().tolist())
        print(f"⚠️  [{origen}] {len(no_reconocidos)} fila(s) con valor de 'Mes' no reconocido "
              f"(se agruparán como '?'): {valores}")
    return mes_num
def _formatear_periodo(mes_num_serie, anio_serie):
    mes_abr  = mes_num_serie.apply(lambda m: MESES_ES[int(m)-1].capitalize()[:3] if m and m > 0 else "?")
    anio_abr = anio_serie.astype(int).astype(str).str[-2:]
    return mes_abr + "-" + anio_abr
def _normalizar_texto(s) -> str:
    s = " ".join(str(s).strip().split())
    s = "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))
    return s.casefold()
def _detectar_columna(df, patron: str):
    patron_norm = _normalizar_texto(patron)
    for c in df.columns:
        if patron_norm in _normalizar_texto(c):
            return c
    return None
def _cargar_imagen_b64(path: str) -> str:
    """Carga una imagen (p.ej. el logo del header) y la codifica en base64
    para incrustarla directamente en el HTML, sin depender de un archivo
    externo que viaje junto al reporte. Si el archivo no existe, retorna
    cadena vacía y el <img> del header simplemente no se muestra (no rompe
    la generación del dashboard)."""
    p = Path(path)
    if not p.exists():
        print(f"⚠️  No se encontró el logo en {path}; se omite del header.")
        return ""
    return base64.b64encode(p.read_bytes()).decode("utf-8")
def formatear_fechas(base: date, es_cierre: bool = False):
    ultimo_dia_dato = base if es_cierre else base - timedelta(days=1)
    fecha_reporte = f"{base.day} de {MESES_ES[base.month-1].capitalize()} de {base.year}"
    dia_semana    = DIAS_ES[ultimo_dia_dato.weekday()].capitalize()
    fecha_info    = f"{dia_semana}, {ultimo_dia_dato.day} de {MESES_ES[ultimo_dia_dato.month-1]} de {ultimo_dia_dato.year}"
    mes_header    = MESES_ES[base.month-1].capitalize()
    return fecha_reporte, fecha_info, mes_header
def periodo_label_actual(base: date) -> str:
    """Devuelve el PeriodoLabel (p.ej. 'Ago-26') correspondiente a FECHA_BASE,
    con el MISMO formato que _formatear_periodo, para poder anclar en el
    front-end las secciones que deben quedar fijas al mes en curso
    (pestaña 'Resumen Mes en Curso') sin depender del segmentador de meses."""
    mes_abr  = MESES_ES[base.month-1].capitalize()[:3]
    anio_abr = str(base.year)[-2:]
    return f"{mes_abr}-{anio_abr}"
def separar_mes_actual_anterior(vmc, fecha_base):
    """A inicios de mes, 'VentasMesCurso' puede traer mezcladas filas del mes
    que ya cerró (p.ej. julio) junto con las del mes nuevo (p.ej. agosto).
    Esta función separa el DataFrame en:
      - vmc_actual: sólo filas cuyo año/mes coincide con FECHA_BASE (mes en curso real)
      - meses_anteriores: lista de dicts (uno por cada mes "sobrante" detectado)
        con la info necesaria para armar una pestaña "Resumen_<Mes>" congelada.
    Si no hay mezcla (todo el mes coincide con fecha_base), meses_anteriores
    queda vacío y el comportamiento es idéntico al de antes.
    """
    fechas = pd.to_datetime(vmc["Fechas de emisión"])
    es_actual = (fechas.dt.year == fecha_base.year) & (fechas.dt.month == fecha_base.month)
    vmc_actual   = vmc[es_actual].copy()
    vmc_anterior = vmc[~es_actual].copy()
    meses_anteriores = []
    if len(vmc_anterior):
        fechas_ant = pd.to_datetime(vmc_anterior["Fechas de emisión"])
        periodos = sorted(set(zip(fechas_ant.dt.year, fechas_ant.dt.month)))
        for anio, mes in periodos:
            label = MESES_ES[mes-1].capitalize()
            slug  = _normalizar_texto(label).replace(" ", "_")
            mask_periodo = (fechas_ant.dt.year == anio) & (fechas_ant.dt.month == mes)
            df_periodo = vmc_anterior[mask_periodo].copy()
            meses_anteriores.append({
                "anio": anio, "mes": mes, "label": label, "slug": slug, "df": df_periodo
            })
            print(f"ℹ️  [MesCurso] Se detectaron {len(df_periodo)} fila(s) de {label} {anio} mezcladas "
                  f"en 'VentasMesCurso'; se reportarán aparte en la pestaña 'Resumen {label}' "
                  f"y se excluyen del cálculo del mes en curso ({MESES_ES[fecha_base.month-1].capitalize()} "
                  f"{fecha_base.year}).")
    return vmc_actual, meses_anteriores
def procesar_mes_curso(vmc, suc, bol_list):
    vmc["ClaveSucursal"] = pd.to_numeric(vmc["ClaveSucursal"], errors="coerce").fillna(0).astype(int)
    vmc = vmc.merge(suc, on="ClaveSucursal", how="left")
    vmc["FechaStr"] = pd.to_datetime(vmc["Fechas de emisión"]).dt.strftime("%Y-%m-%d")
    vmc_v = vmc[~vmc["Artículo"].isin(bol_list)]
    agg_v = vmc_v.groupby(["FechaStr","NombreSucursal"]).agg(
        unidades  =("Unidades","sum"),
        ventas    =("Importe c/Desc","sum"),
        utilidad  =("Utilidad","sum"),
    ).reset_index()
    agg_t = vmc.groupby(["FechaStr","NombreSucursal"]).agg(
        tickets=("Movimiento","nunique")
    ).reset_index()
    agg = agg_t.merge(agg_v, on=["FechaStr","NombreSucursal"], how="left")
    for c in ["unidades","ventas","utilidad"]:
        agg[c] = agg[c].fillna(0).round(2)
    agg["tickets"] = agg["tickets"].fillna(0).astype(int)
    agg["margen"]  = _margen_seguro(agg["utilidad"], agg["ventas"])
    return agg
def procesar_lineas(vm, art, suc, bol_list):
    vm = vm.copy()
    vm["ClaveSucursal"] = pd.to_numeric(vm["ClaveSucursal"], errors="coerce").fillna(0).astype(int)
    vm = vm[~vm["Artículo"].isin(bol_list)]
    art_clean = art.drop_duplicates(subset=["Artículo"])
    suc_clean = suc.drop_duplicates(subset=["ClaveSucursal"])
    vm_merged = vm.merge(art_clean, on="Artículo", how="left").merge(suc_clean, on="ClaveSucursal", how="left")
    vm_merged["Línea"]          = vm_merged["Línea"].fillna("NO ASIGNADO")
    vm_merged["NombreSucursal"] = vm_merged["NombreSucursal"].fillna("OTRO")
    vm_merged["MesNum"] = _mapear_mes(vm_merged["Mes"], "VentasMensuales (Resumen por Línea)").fillna(0).astype(int)
    vm_merged["Año"]    = pd.to_numeric(vm_merged["Año"], errors="coerce").fillna(0).astype(int)
    vm_merged["PeriodoLabel"] = _formatear_periodo(vm_merged["MesNum"], vm_merged["Año"])
    linea_agg = vm_merged.groupby(["NombreSucursal","Línea","PeriodoLabel"]).agg(
        ventas    =("Importe c/Desc","sum"),
        utilidad  =("Utilidad","sum"),
        costo     =("Costo","sum"),
        unidades  =("Unidades","sum"),
    ).reset_index()
    for c in ["ventas","utilidad","costo"]:
        linea_agg[c] = linea_agg[c].round(2)
    linea_agg["margen"] = _margen_seguro(linea_agg["utilidad"], linea_agg["ventas"])
    return linea_agg
def procesar_historico(vm, tkt, suc, bol_list):
    vm["ClaveSucursal"]  = pd.to_numeric(vm["ClaveSucursal"],  errors="coerce").fillna(0).astype(int)
    tkt["ClaveSucursal"] = pd.to_numeric(tkt["ClaveSucursal"], errors="coerce").fillna(0).astype(int)
    vm = vm[~vm["Artículo"].isin(bol_list)].copy()
    vm["MesNum"] = _mapear_mes(vm["Mes"], "VentasMensuales")
    agg_v = vm.groupby(["Año","MesNum","ClaveSucursal"]).agg(
        unidades  =("Unidades","sum"),
        ventas    =("Importe c/Desc","sum"),
        utilidad  =("Utilidad","sum"),
    ).reset_index()
    tkt["MesNum"] = _mapear_mes(tkt["Mes"], "TicketsMensuales")
    mask_consolidado = tkt["ClaveSucursal"].isin(CLAVE_CONSOLIDADA)
    tkt_consolidado = tkt[mask_consolidado].copy()
    tkt = tkt[~mask_consolidado].copy()
    tickets_consolidados = {}
    if len(tkt_consolidado):
        cons = tkt_consolidado.groupby(["Año","MesNum"]).agg(
            tickets=("Tickets","sum")
        ).reset_index()
        cons["Año"]    = pd.to_numeric(cons["Año"], errors="coerce").fillna(0).astype(int)
        cons["MesNum"] = cons["MesNum"].fillna(0).astype(int)
        cons["PeriodoLabel"] = _formatear_periodo(cons["MesNum"], cons["Año"])
        tickets_consolidados = {
            str(r.PeriodoLabel): int(r.tickets) for r in cons.itertuples()
        }
        print(f"ℹ️  [Histórico] {len(tickets_consolidados)} período(s) con tickets consolidados bajo la(s) "
              f"clave(s) {CLAVE_CONSOLIDADA}: {', '.join(sorted(tickets_consolidados))}. Se mostrarán como "
              f"total del mes con la nota 'Sin detalle por canal de venta'.")
    agg_t = tkt.groupby(["Año","MesNum","ClaveSucursal"]).agg(
        tickets=("Tickets","sum")
    ).reset_index()
    agg = agg_v.merge(agg_t, on=["Año","MesNum","ClaveSucursal"], how="outer")
    for c in ["unidades","ventas","utilidad"]:
        agg[c] = agg[c].fillna(0).round(2)
    agg["Año"]    = agg["Año"].fillna(0).astype(int)
    agg["MesNum"] = agg["MesNum"].fillna(0).astype(int)
    suc_clean = suc.drop_duplicates(subset=["ClaveSucursal"])
    agg = agg.merge(suc_clean, on="ClaveSucursal", how="left")
    agg["NombreSucursal"] = agg["NombreSucursal"].fillna("OTRO")
    agg["tickets"] = agg["tickets"].fillna(0).astype(int)
    agg["margen"] = _margen_seguro(agg["utilidad"], agg["ventas"])
    agg["PeriodoLabel"] = _formatear_periodo(agg["MesNum"], agg["Año"])
    agg = agg.sort_values(["Año","MesNum"]).reset_index(drop=True)
    return agg, tickets_consolidados
def procesar_top_articulos(vm, art_dim, suc, bol_list):
    vm = vm.copy()
    vm["ClaveSucursal"] = pd.to_numeric(vm["ClaveSucursal"], errors="coerce").fillna(0).astype(int)
    vm = vm[~vm["Artículo"].isin(bol_list)]
    suc_clean = suc.drop_duplicates(subset=["ClaveSucursal"])
    vm = vm.merge(suc_clean, on="ClaveSucursal", how="left")
    vm["NombreSucursal"] = vm["NombreSucursal"].fillna("OTRO")
    art_clean = art_dim.drop_duplicates(subset=["Artículo"])
    vm = vm.merge(art_clean[["Artículo","Descripción","Fabricante"]], on="Artículo", how="left")
    vm["Descripción"] = vm["Descripción"].fillna("SIN DESCRIPCIÓN")
    vm["Fabricante"]  = vm["Fabricante"].fillna("SIN FABRICANTE")
    vm["MesNum"] = _mapear_mes(vm["Mes"], "VentasMensuales (Top Artículos)").fillna(0).astype(int)
    vm["Año"]    = pd.to_numeric(vm["Año"], errors="coerce").fillna(0).astype(int)
    vm["PeriodoLabel"] = _formatear_periodo(vm["MesNum"], vm["Año"])
    agg = vm.groupby(["Artículo","Descripción","Fabricante","NombreSucursal","PeriodoLabel"]).agg(
        unidades =("Unidades","sum"),
        ventas   =("Importe c/Desc","sum"),
        utilidad =("Utilidad","sum"),
    ).reset_index()
    for c in ["ventas","utilidad"]:
        agg[c] = agg[c].round(2)
    return agg
def procesar_lineas_categoria(vm, art_dim, suc, bol_list):
    vm = vm.copy()
    vm["ClaveSucursal"] = pd.to_numeric(vm["ClaveSucursal"], errors="coerce").fillna(0).astype(int)
    vm = vm[~vm["Artículo"].isin(bol_list)]
    suc_clean = suc.drop_duplicates(subset=["ClaveSucursal"])
    vm = vm.merge(suc_clean, on="ClaveSucursal", how="left")
    vm["NombreSucursal"] = vm["NombreSucursal"].fillna("OTRO")
    art_clean = art_dim.drop_duplicates(subset=["Artículo"])
    vm = vm.merge(art_clean[["Artículo","Línea","Categoría"]], on="Artículo", how="left")
    vm["Línea"]     = vm["Línea"].fillna("NO ASIGNADO")
    vm["Categoría"] = vm["Categoría"].fillna("SIN CATEGORÍA")
    vm["MesNum"] = _mapear_mes(vm["Mes"], "VentasMensuales (Líneas y Categorías)").fillna(0).astype(int)
    vm["Año"]    = pd.to_numeric(vm["Año"], errors="coerce").fillna(0).astype(int)
    vm["PeriodoLabel"] = _formatear_periodo(vm["MesNum"], vm["Año"])
    agg = vm.groupby(["Línea","Categoría","NombreSucursal","PeriodoLabel"]).agg(
        unidades =("Unidades","sum"),
        ventas   =("Importe c/Desc","sum"),
        utilidad =("Utilidad","sum"),
    ).reset_index()
    for c in ["ventas","utilidad"]:
        agg[c] = agg[c].round(2)
    return agg
def procesar_fabricantes(vm, art_dim, suc, bol_list):
    vm = vm.copy()
    vm["ClaveSucursal"] = pd.to_numeric(vm["ClaveSucursal"], errors="coerce").fillna(0).astype(int)
    vm = vm[~vm["Artículo"].isin(bol_list)]
    suc_clean = suc.drop_duplicates(subset=["ClaveSucursal"])
    vm = vm.merge(suc_clean, on="ClaveSucursal", how="left")
    vm["NombreSucursal"] = vm["NombreSucursal"].fillna("OTRO")
    art_clean = art_dim.drop_duplicates(subset=["Artículo"])
    vm = vm.merge(art_clean[["Artículo","Fabricante"]], on="Artículo", how="left")
    vm["Fabricante"] = vm["Fabricante"].fillna("SIN FABRICANTE")
    vm["MesNum"] = _mapear_mes(vm["Mes"], "VentasMensuales (Fabricantes)").fillna(0).astype(int)
    vm["Año"]    = pd.to_numeric(vm["Año"], errors="coerce").fillna(0).astype(int)
    vm["PeriodoLabel"] = _formatear_periodo(vm["MesNum"], vm["Año"])
    agg = vm.groupby(["Fabricante","NombreSucursal","PeriodoLabel"]).agg(
        unidades =("Unidades","sum"),
        ventas   =("Importe c/Desc","sum"),
        utilidad =("Utilidad","sum"),
    ).reset_index()
    for c in ["ventas","utilidad"]:
        agg[c] = agg[c].round(2)
    return agg
def procesar_pronostico(agg, objetivos, suc, fecha_base, es_cierre=False):
    """Pronóstico de cierre de mes por sucursal.

    Skin Love NO maneja presupuestos de venta, así que aquí no hay ninguna
    lógica de presupuesto, cumplimiento ni alcance contra objetivo: el
    pronóstico se calcula únicamente proyectando el ritmo de venta observado
    (ventas acumuladas ÷ días transcurridos × días operativos del mes).

    De la hoja 'ObjetivosVentas' sólo se lee la columna 'FechaApertura', que
    sirve para no castigar a una sucursal recién abierta: sus días
    transcurridos se cuentan desde su apertura, no desde el día 1 del mes.
    Si no hay FechaApertura registrada, se infiere con la fecha de su primera
    venta del mes.

    Se excluyen las sucursales sin venta en el mes en curso: un pronóstico de
    $0 no aporta información y sólo alarga la tabla.
    """
    dias_mes = calendar.monthrange(fecha_base.year, fecha_base.month)[1]
    if es_cierre:
        fecha_max_global = pd.Timestamp(fecha_base)
    else:
        fecha_max_global = pd.Timestamp(fecha_base) - pd.Timedelta(days=1)
    primer_dia_mes = pd.Timestamp(year=fecha_base.year, month=fecha_base.month, day=1)
    fin_mes = pd.Timestamp(year=fecha_base.year, month=fecha_base.month, day=dias_mes)
    resumen_ventas = agg.groupby("NombreSucursal").agg(
        unidadesActual=("unidades", "sum"),
        ventasActual=("ventas", "sum"),
        utilidadActual=("utilidad", "sum"),
        fechaMin=("FechaStr", "min"),
    ).reset_index()
    resumen_ventas["fechaMinDt"] = pd.to_datetime(resumen_ventas["fechaMin"])
    suc_clean = suc.drop_duplicates(subset=["NombreSucursal"])[["ClaveSucursal", "NombreSucursal"]].copy()
    suc_clean["NombreSucursal"] = suc_clean["NombreSucursal"].astype(str).str.strip()
    resumen_ventas["NombreSucursal"] = resumen_ventas["NombreSucursal"].astype(str).str.strip()
    resumen = suc_clean.merge(resumen_ventas, on="NombreSucursal", how="left")
    for c in ["unidadesActual", "ventasActual", "utilidadActual"]:
        resumen[c] = resumen[c].fillna(0)
    resumen["margen"] = _margen_seguro(resumen["utilidadActual"], resumen["ventasActual"])
    col_suc = _detectar_columna(objetivos, "sucursal")
    col_ape = _detectar_columna(objetivos, "apertura")
    resumen["FechaApertura"] = pd.NaT
    if col_suc is None or col_ape is None:
        print("⚠️  [Pronóstico] No se encontró columna 'Sucursal' y/o 'FechaApertura' en la hoja "
              f"'ObjetivosVentas' (columnas disponibles: {list(objetivos.columns)}). Se usará el método "
              "de inferencia (primera venta del mes) para calcular los días de todas las sucursales.")
    else:
        obj_clean = objetivos[[col_suc, col_ape]].rename(
            columns={col_suc: "SucursalObjetivo", col_ape: "FechaApertura"}
        ).dropna(subset=["SucursalObjetivo"]).copy()
        obj_clean["FechaApertura"] = pd.to_datetime(obj_clean["FechaApertura"], dayfirst=True, errors="coerce")
        claves_num = pd.to_numeric(obj_clean["SucursalObjetivo"], errors="coerce")
        usa_clave = claves_num.notna().mean() >= 0.5 if len(obj_clean) else False
        if usa_clave:
            obj_clean["ClaveSucursal"] = claves_num.fillna(0).astype(int)
            obj_clean = obj_clean.drop_duplicates(subset=["ClaveSucursal"])
            resumen = resumen.drop(columns=["FechaApertura"]).merge(
                obj_clean[["ClaveSucursal", "FechaApertura"]], on="ClaveSucursal", how="left"
            )
            keys_dim  = set(resumen["ClaveSucursal"])
            keys_obj  = set(obj_clean["ClaveSucursal"])
            huerfanos = sorted(keys_obj - keys_dim)
            modo = "ClaveSucursal (código numérico)"
        else:
            resumen["_key"]   = resumen["NombreSucursal"].apply(_normalizar_texto)
            obj_clean["_key"] = obj_clean["SucursalObjetivo"].apply(_normalizar_texto)
            obj_clean = obj_clean.drop_duplicates(subset=["_key"])
            resumen = resumen.drop(columns=["FechaApertura"]).merge(
                obj_clean[["_key", "FechaApertura"]], on="_key", how="left"
            )
            keys_dim = set(resumen["_key"])
            keys_obj = set(obj_clean["_key"])
            huerfanos = sorted(
                obj_clean.loc[obj_clean["_key"].isin(keys_obj - keys_dim), "SucursalObjetivo"].unique().tolist()
            )
            resumen = resumen.drop(columns=["_key"])
            modo = "nombre de sucursal (texto normalizado)"
        if huerfanos:
            print(f"⚠️  [Pronóstico] Cruce por {modo}. {len(huerfanos)} valor(es) en 'ObjetivosVentas' NO "
                  f"encontrados en la dimensión de sucursales: {huerfanos}")
    antes = len(resumen)
    sin_movimiento_mask = (resumen["ventasActual"] == 0) & (resumen["unidadesActual"] == 0)
    excluidas = sorted(resumen.loc[sin_movimiento_mask, "NombreSucursal"].tolist())
    resumen = resumen[~sin_movimiento_mask].copy()
    if excluidas:
        print(f"ℹ️  [Pronóstico] {len(excluidas)} de {antes} sucursal(es) se excluyeron por no tener "
              f"movimiento registrado en el mes en curso: {excluidas}")
    usa_inferencia = sorted(
        resumen.loc[resumen["FechaApertura"].isna(), "NombreSucursal"].tolist()
    )
    if usa_inferencia:
        print(f"ℹ️  [Pronóstico] {len(usa_inferencia)} sucursal(es) sin 'FechaApertura' registrada: se usó "
              f"el método de inferencia (primera venta del mes) para calcular sus días: {usa_inferencia}")
    def _calcular_dias(row):
        apertura = row["FechaApertura"]
        if pd.isna(apertura):
            inicio = row["fechaMinDt"] if pd.notna(row["fechaMinDt"]) else primer_dia_mes
        else:
            inicio = apertura
        if inicio < primer_dia_mes:
            inicio = primer_dia_mes
        dias_operativos = max(0, (fin_mes - inicio).days + 1)
        dias_transcurridos = max(0, (fecha_max_global - inicio).days + 1)
        return pd.Series({"diasOperativosMes": dias_operativos, "diasTranscurridos": dias_transcurridos})
    if resumen.empty:
        resumen["diasOperativosMes"] = pd.Series(dtype="int64")
        resumen["diasTranscurridos"] = pd.Series(dtype="int64")
    else:
        resumen[["diasOperativosMes", "diasTranscurridos"]] = resumen.apply(_calcular_dias, axis=1)
    resumen["pronostico"] = np.where(
        resumen["diasTranscurridos"] > 0,
        resumen["ventasActual"] / resumen["diasTranscurridos"] * resumen["diasOperativosMes"],
        0.0,
    )
    for c in ["unidadesActual", "ventasActual", "utilidadActual", "pronostico"]:
        resumen[c] = resumen[c].round(2)
    sin_pronostico_mask = resumen["ventasActual"] <= 0
    if sin_pronostico_mask.any():
        sin_pron = sorted(resumen.loc[sin_pronostico_mask, "NombreSucursal"].tolist())
        print(f"ℹ️  [Pronóstico] {len(sin_pron)} canal(es) con venta acumulada no positiva: se incluyen "
              f"en la tabla sin dato de pronóstico: {sin_pron}")
    resumen["pronostico"] = resumen["pronostico"].astype(object).where(~sin_pronostico_mask, None)
    resumen = resumen.sort_values("ClaveSucursal").reset_index(drop=True)
    return resumen[["ClaveSucursal", "NombreSucursal", "unidadesActual", "ventasActual", "utilidadActual",
                     "margen", "pronostico", "diasTranscurridos", "diasOperativosMes"]]
def ordenar_sucursales_por_apertura(suc: pd.DataFrame, objetivos: pd.DataFrame) -> list:
    """Determina el orden de los botones del segmentador de sucursales:
    de la FechaApertura más antigua a la más reciente (columna 'FechaApertura'
    en la hoja 'ObjetivosVentas', la misma que usa procesar_pronostico para
    el pronóstico). Las sucursales sin FechaApertura registrada se colocan al
    final, ordenadas por ClaveSucursal (ID) ascendente.
    Es una función independiente y autocontenida (no comparte estado con
    procesar_pronostico) para no arriesgar esa lógica ya probada.
    """
    suc_clean = suc.drop_duplicates(subset=["NombreSucursal"])[["ClaveSucursal", "NombreSucursal"]].copy()
    suc_clean["NombreSucursal"] = suc_clean["NombreSucursal"].astype(str).str.strip()
    suc_clean["FechaApertura"] = pd.NaT
    col_suc = _detectar_columna(objetivos, "sucursal")
    col_ape = _detectar_columna(objetivos, "apertura")
    if col_suc is None or col_ape is None:
        print("⚠️  [Segmentador Sucursales] No se encontró columna 'Sucursal' y/o 'FechaApertura' "
              "en 'ObjetivosVentas'; las sucursales se ordenarán únicamente por ClaveSucursal (ID).")
    else:
        obj_clean = objetivos[[col_suc, col_ape]].rename(
            columns={col_suc: "SucursalObjetivo", col_ape: "FechaApertura"}
        ).dropna(subset=["SucursalObjetivo"]).copy()
        obj_clean["FechaApertura"] = pd.to_datetime(obj_clean["FechaApertura"], dayfirst=True, errors="coerce")
        claves_num = pd.to_numeric(obj_clean["SucursalObjetivo"], errors="coerce")
        usa_clave = claves_num.notna().mean() >= 0.5 if len(obj_clean) else False
        if usa_clave:
            obj_clean["ClaveSucursal"] = claves_num.fillna(0).astype(int)
            obj_clean = obj_clean.drop_duplicates(subset=["ClaveSucursal"])
            suc_clean = suc_clean.drop(columns=["FechaApertura"]).merge(
                obj_clean[["ClaveSucursal", "FechaApertura"]], on="ClaveSucursal", how="left"
            )
        else:
            suc_clean["_key"] = suc_clean["NombreSucursal"].apply(_normalizar_texto)
            obj_clean["_key"] = obj_clean["SucursalObjetivo"].apply(_normalizar_texto)
            obj_clean = obj_clean.drop_duplicates(subset=["_key"])
            suc_clean = suc_clean.drop(columns=["FechaApertura"]).merge(
                obj_clean[["_key", "FechaApertura"]], on="_key", how="left"
            )
            suc_clean = suc_clean.drop(columns=["_key"])
    sin_fecha_n = suc_clean["FechaApertura"].isna().sum()
    if sin_fecha_n:
        nombres_sin_fecha = sorted(suc_clean.loc[suc_clean["FechaApertura"].isna(), "NombreSucursal"].tolist())
        print(f"ℹ️  [Segmentador Sucursales] {sin_fecha_n} sucursal(es) sin FechaApertura registrada; "
              f"se colocarán al final, ordenadas por ClaveSucursal: {nombres_sin_fecha}")
    con_fecha = suc_clean[suc_clean["FechaApertura"].notna()].sort_values("FechaApertura")
    sin_fecha = suc_clean[suc_clean["FechaApertura"].isna()].sort_values("ClaveSucursal")
    orden_final = pd.concat([con_fecha, sin_fecha], ignore_index=True)
    return orden_final["NombreSucursal"].tolist()
def generar_html(agg, linea_agg, historico_agg, top_art_agg, lineas_cat_agg, fabricantes_agg,
                 pronostico_agg, lista_sucursales, fecha_reporte, fecha_info, mes_header,
                 current_period_label, tickets_consolidados=None, resumenes_anteriores=None):
    resumenes_anteriores = resumenes_anteriores or []
    tabs_nav_extra = ""
    tabs_content_extra = ""
    resumenes_ant_dict = {}
    for r in resumenes_anteriores:
        slug, label, anio = r["slug"], r["label"], r["anio"]
        tabid = f"resumen_{slug}"
        tabs_nav_extra += (f'\n  <button class="tab-nav-btn" id="tabnav-{tabid}" '
                            f'onclick="switchTab(\'{tabid}\')">Resumen {label}</button>')
        tabs_content_extra += f"""
<div id="tab-{tabid}" class="tab-content">
<div class="main">
  <div class="tc">
    <div class="card-head">
      <div><div class="card-title">Ventas Diarias · {label} {anio}</div><div class="card-sub">Mes cerrado · detalle completo por día</div></div>
      <span class="note-bol">Canales de venta según selección</span>
    </div>
    <div class="table-scale-wrap">
    <table>
      <thead><tr><th>Fecha</th><th>Día</th><th class="r">Unidades</th><th class="r">Ventas $</th><th class="r">Utilidad</th><th class="r">Margen</th><th class="r">Tickets</th></tr></thead>
      <tbody id="tabla-resumen-{slug}"></tbody>
    </table>
    </div>
  </div>
  <div class="charts-row">
    <div class="cc"><div class="card-head" style="margin-bottom:.4rem"><div><div class="card-title">Ventas $</div><div class="card-sub">Volumen diario · Canales de venta seleccionados</div></div></div><div class="cw"><canvas id="chart-resumen-{slug}-ventas"></canvas></div></div>
    <div class="cc"><div class="card-head" style="margin-bottom:.4rem"><div><div class="card-title">No. de Tickets</div><div class="card-sub">Volumen diario · Canales de venta seleccionados</div></div></div><div class="cw"><canvas id="chart-resumen-{slug}-tickets"></canvas></div></div>
  </div>
</div>
</div>"""
        resumenes_ant_dict[slug] = {"label": label, "anio": anio, "data": r["agg"].to_dict("records")}
    resumenes_ant_json = json.dumps(resumenes_ant_dict, ensure_ascii=False)
    data_json         = json.dumps(agg.to_dict("records"),           ensure_ascii=False)
    linea_json        = json.dumps(linea_agg.to_dict("records"),     ensure_ascii=False)
    historico_json    = json.dumps(historico_agg.to_dict("records"), ensure_ascii=False)
    top_art_json      = json.dumps(top_art_agg.to_dict("records"),   ensure_ascii=False)
    lineas_cat_json   = json.dumps(lineas_cat_agg.to_dict("records"),ensure_ascii=False)
    fabricantes_json  = json.dumps(fabricantes_agg.to_dict("records"),ensure_ascii=False)
    pronostico_json   = json.dumps(pronostico_agg.to_dict("records"),ensure_ascii=False)
    sucursales_json   = json.dumps(lista_sucursales,                  ensure_ascii=False)
    current_period_json = json.dumps(current_period_label,            ensure_ascii=False)
    tickets_cons_json   = json.dumps(tickets_consolidados or {},       ensure_ascii=False)
    logo_b64 = _cargar_imagen_b64(LOGO_PATH)
    periodos_unicos = (
        historico_agg[["Año","MesNum","PeriodoLabel"]]
        .drop_duplicates()
        .sort_values(["Año","MesNum"])["PeriodoLabel"]
        .tolist()
    )
    periodos_json = json.dumps(periodos_unicos, ensure_ascii=False)
    html_template = r"""<!DOCTYPE html>
<html lang="es">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Skin Love · Análisis de Ventas</title>
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.2.0"></script>
<script src="https://cdn.jsdelivr.net/npm/chartjs-chart-treemap@3/dist/chartjs-chart-treemap.min.js"></script>
<style>
*{box-sizing:border-box;margin:0;padding:0}
html{font-size:clamp(13px, 4vw, 16px)}
body{font-family:'Segoe UI',system-ui,sans-serif;background:#fcfcfc;color:#212121;min-height:100vh}
.sticky-header{position:sticky;top:0;z-index:1000;isolation:isolate}
header{background:#808080;padding:1.2rem 2rem;display:grid;grid-template-columns:1fr auto 1fr;column-gap:1rem;align-items:center;box-shadow:0 3px 20px rgba(0,0,0,0.12)}
.header-left{min-width:0}
.header-left h1{font-size:1.15rem;font-weight:700;color:#fff;letter-spacing:.03em}
.header-left p{font-size:.72rem;color:#f2f2f2;margin-top:4px;letter-spacing:0.02em}
.header-right{display:flex;flex-direction:column;align-items:flex-end;gap:.35rem;min-width:0}
.hbadge{background:#ffffff1f;border:1px solid #ffffff40;color:#fff;font-size:.68rem;font-weight:600;padding:4px 14px;border-radius:20px;letter-spacing:.05em}
.hdate{font-size:.72rem;color:#f2f2f2;font-weight:500}
.header-logo{display:flex;align-items:center;justify-content:center;min-width:0}
.header-logo img{height:6.25rem;width:auto;display:block}
.filter-bar{background:#fff;border-bottom:1px solid #ebebeb;padding:.75rem 2rem;display:flex;gap:.8rem;flex-wrap:wrap;align-items:center;justify-content:center;box-shadow:0 2px 4px rgba(0,0,0,0.02)}
.filter-bar + .filter-bar{border-top:1px solid #f2f2f2;box-shadow:none}
.suc-label{font-size:.67rem;font-weight:700;text-transform:uppercase;letter-spacing:.1em;color:#5A5A5A;margin-right:.25rem;white-space:nowrap}
.suc-hint{font-size:.62rem;color:#a3a3a3;font-weight:500;white-space:nowrap;margin-right:.4rem}
.suc-btns{display:flex;gap:6px;flex-wrap:wrap;justify-content:center}
.suc-btn{display:inline-flex;align-items:center;padding:6px 13px;border-radius:20px;border:1px solid #dedede;background:#f7f7f7;color:#6b6b6b;font-size:.72rem;font-weight:600;cursor:pointer;transition:all .2s;font-family:inherit;white-space:nowrap}
.suc-btn.active{background:#5A5A5A !important;border-color:#5A5A5A !important;color:#fff !important;font-weight:700;box-shadow:0 3px 10px rgba(0,0,0,0.15)}
.mes-btn{display:inline-flex;align-items:center;padding:5px 11px;border-radius:20px;border:1px solid #dedede;background:#f7f7f7;color:#6b6b6b;font-size:.69rem;font-weight:600;cursor:pointer;transition:all .2s;font-family:inherit;white-space:nowrap}
.mes-btn.active{background:#3D3D3D !important;border-color:#3D3D3D !important;color:#fff !important;font-weight:700}
.suc-btn.oculto,.mes-btn.oculto{display:none}
.suc-sep{width:1px;height:20px;background:#e0e0e0;margin:0 .1rem;flex-shrink:0}
.btn-all{padding:6px 13px;border-radius:20px;border:1px solid #5A5A5A;background:#fff;color:#5A5A5A;font-size:.7rem;font-weight:700;cursor:pointer;font-family:inherit;transition:all .2s;white-space:nowrap}
.btn-all.dark{border-color:#3D3D3D;color:#3D3D3D}
.tabs-nav{background:#fff;border-bottom:2px solid #e6e6e6;padding:0 2rem;display:flex;justify-content:center;gap:0;flex-wrap:wrap}
.tab-nav-btn{padding:.85rem 1.6rem;font-size:.82rem;font-weight:700;color:#858585;cursor:pointer;border:none;background:none;font-family:inherit;border-bottom:3px solid transparent;margin-bottom:-2px;transition:all .2s;letter-spacing:.02em}
.tab-nav-btn.active{color:#3D3D3D;border-bottom-color:#3D3D3D}
.tab-nav-btn:hover:not(.active){color:#5A5A5A;border-bottom-color:#dedede}
.tab-content{display:none}
.tab-content.active{display:block}
.main{padding:1.5rem 2rem 3rem;max-width:1450px;margin:0 auto}
.kpi-grid{display:flex;flex-wrap:wrap;justify-content:center;gap:.9rem;margin-bottom:1.5rem}
.kpi{background:#fff;border-radius:12px;padding:1.1rem 1.2rem;border:1px solid #e6e6e6;position:relative;overflow:hidden;box-shadow:0 2px 6px rgba(0,0,0,0.015);flex:1 1 170px;max-width:215px;min-width:0;text-align:center}
.kpi::before{content:'';position:absolute;top:0;left:0;right:0;height:3px;background:var(--ac,#808080)}
.kpi-label{font-size:.63rem;font-weight:700;text-transform:uppercase;letter-spacing:.09em;color:#858585;margin-bottom:6px}
.kpi-value{font-size:1.4rem;font-weight:700;color:#212121;line-height:1}
.kpi-value.brand{color:#5A5A5A}
.kpi-value.gold{color:#aa7300}
.kpi-sub{font-size:.67rem;color:#858585;margin-top:5px}
.kpi-note{font-size:.68rem;color:#858585;margin:-.6rem 0 1rem;font-style:italic}
.tc{background:#fff;border-radius:12px;padding:1.2rem 1.6rem;border:1px solid #e6e6e6;margin-bottom:.9rem;box-shadow:0 2px 6px rgba(0,0,0,0.015)}
.card-head{display:flex;align-items:center;justify-content:space-between;margin-bottom:1.1rem;gap:1rem;flex-wrap:wrap}
.card-title{font-size:.88rem;font-weight:700;color:#212121;letter-spacing:0.01em}
.card-sub{font-size:.7rem;color:#858585;margin-top:2px}
.note-bol{font-size:.67rem;background:#fafafa;color:#5A5A5A;border:1px solid #e0e0e0;padding:4px 10px;border-radius:6px;white-space:nowrap;font-weight:500}
.metric-tabs{display:flex;gap:4px;background:#f2f2f2;padding:3px;border-radius:8px;border:1px solid #e4e4e4}
.tab-btn{background:none;border:none;padding:5px 12px;font-size:.72rem;font-weight:600;color:#6b6b6b;cursor:pointer;border-radius:6px;font-family:inherit;transition:all .15s}
.tab-btn.active{background:#fff;color:#3D3D3D;box-shadow:0 2px 5px rgba(0,0,0,0.08);font-weight:700}
table{width:100%;border-collapse:collapse;font-size:.78rem;min-width:100%}
thead tr{background:#fcfcfc;border-bottom:2px solid #e6e6e6}
thead th{text-align:left;padding:9px 11px;font-size:.63rem;font-weight:700;text-transform:uppercase;letter-spacing:.09em;color:#5A5A5A;white-space:nowrap}
thead th.r{text-align:right}
tbody tr{border-bottom:1px solid #f4f4f4;transition:background .15s}
tbody tr:last-child{border-bottom:none}
tbody tr:hover{background:#fafafa}
tbody tr.total-row{background:linear-gradient(90deg,#fafafa,#fff);border-top:2px solid #c4c4c4}
tbody tr.total-row td{font-weight:700;color:#3D3D3D}
td{padding:9px 11px;color:#4d4d4d}
td.r{text-align:right;font-variant-numeric:tabular-nums}
td.date{font-weight:600;color:#212121;white-space:nowrap}
td.dayname{font-size:.68rem;color:#858585;white-space:nowrap}
td.art-code{font-family:'Consolas',monospace;font-size:.72rem;color:#5A5A5A;font-weight:700;white-space:nowrap}
td.art-desc{font-size:.75rem;color:#4d4d4d;line-height:1.4;word-break:break-word;min-width:200px}
td.art-fab{font-size:.72rem;color:#6b6b6b;word-break:break-word;min-width:120px}
td.rank{font-size:.78rem;font-weight:800;color:#858585;width:32px;text-align:center}
td.rank.top3{color:#5A5A5A}
.pill{display:inline-block;padding:2px 9px;border-radius:20px;font-size:.65rem;font-weight:600;letter-spacing:0.02em}
.pill.hi{background:#e3f7ed;color:#176440}
.pill.mi{background:#fff3db;color:#805200}
.pill.lo{background:#fbe4e4;color:#a12727}
.charts-row{display:grid;grid-template-columns:repeat(auto-fit,minmax(320px,1fr));gap:.9rem}
.cc{background:#fff;border-radius:12px;padding:1.2rem 1.6rem;border:1px solid #e6e6e6;box-shadow:0 2px 6px rgba(0,0,0,0.015)}
.lines-container{display:grid;grid-template-columns:minmax(0,3fr) minmax(0,2fr);gap:.9rem;margin-top:1.5rem;align-items:start}
.lines-table-box,.lines-chart-box{background:#fff;border-radius:12px;padding:1.2rem 1.6rem;border:1px solid #e6e6e6;box-shadow:0 2px 6px rgba(0,0,0,0.015);display:flex;flex-direction:column;min-width:0}
.hist-card{background:#fff;border-radius:12px;padding:1.2rem 1.6rem;border:1px solid #e6e6e6;box-shadow:0 2px 6px rgba(0,0,0,0.015)}
.cw{position:relative;height:265px}
.cw.bars-horizontal{min-height:220px}
.cw.hist-chart{position:relative;height:360px;width:100%}
.cw.bubble-chart{position:relative;height:560px;width:100%}
.empty{text-align:center;padding:2.5rem;color:#858585;font-size:.82rem}
.section-divider{border:none;border-top:2px solid #ebebeb;margin:2rem 0}
.top-header-card{background:#fff;border-radius:12px;padding:1rem 1.6rem;border:1px solid #e6e6e6;box-shadow:0 2px 6px rgba(0,0,0,0.015);margin-bottom:1rem}
.top-layout{display:block}
.top-table-box{background:#fff;border-radius:12px;padding:1.2rem 1.6rem;border:1px solid #e6e6e6;box-shadow:0 2px 6px rgba(0,0,0,0.015);margin-bottom:1rem}
.top-bubble-box{background:#fff;border-radius:12px;padding:1.2rem 1.6rem;border:1px solid #e6e6e6;box-shadow:0 2px 6px rgba(0,0,0,0.015);display:flex;flex-direction:column;min-height:0}
.badge-rank{display:inline-flex;align-items:center;justify-content:center;width:22px;height:22px;border-radius:50%;font-size:.68rem;font-weight:800;background:#f0f0f0;color:#858585}
.badge-rank.r1{background:#3D3D3D;color:#fff}
.badge-rank.r2{background:#6B6B6B;color:#fff}
.badge-rank.r3{background:#A8A8A8;color:#262626}
.top-kpi-row{display:grid;grid-template-columns:repeat(3,1fr);gap:.75rem;margin-bottom:1rem}
.top-kpi{background:#fff;border-radius:10px;padding:.9rem 1.1rem;border:1px solid #e6e6e6;position:relative;overflow:hidden;box-shadow:0 2px 6px rgba(0,0,0,0.015);text-align:center}
.top-kpi::before{content:'';position:absolute;top:0;left:0;right:0;height:3px;background:var(--ac,#808080)}
.top-kpi .kpi-label{font-size:.6rem}
.top-kpi .kpi-value{font-size:1.1rem}
.legend-gradient{display:inline-flex;align-items:center;gap:6px;font-size:.67rem;color:#6b6b6b;white-space:nowrap}
.legend-bar{width:70px;height:8px;border-radius:4px;background:linear-gradient(90deg,#d9d9d9,#3D3D3D)}
tr.lc-parent{cursor:pointer}
tr.lc-parent:hover{background:#fafafa}
tr.lc-child td{background:#fafafa;color:#6b6b6b}
tr.lc-child td:nth-child(2){padding-left:26px}
.lc-icon{display:inline-block;width:14px;color:#5A5A5A;font-weight:700}
.fq-badge{display:inline-block;padding:2px 10px;border-radius:20px;font-size:.65rem;font-weight:700;letter-spacing:.02em;white-space:nowrap}
tr.fq-parent{cursor:pointer}
tr.fq-parent:hover td{filter:brightness(0.97)}
tr.fq-parent td{padding-top:10px;padding-bottom:10px}
.fab-bar-cell{position:relative}
.fab-bar-bg{position:absolute;right:0;top:3px;bottom:3px;border-radius:4px;z-index:0}
.fab-bar-cell b{position:relative;z-index:1}
.fab-rank-num{font-size:.68rem;color:#858585;font-weight:700}
.table-scale-wrap{
  width:100%;
  overflow:hidden;
}
.table-scale-wrap table{
  transform-origin: top left;
}
@media (max-width: 768px){
  .lines-container{grid-template-columns:1fr}
}
@media (max-width: 640px){
  .main{padding:1.2rem 1rem 2.5rem}
  .kpi-grid{gap:.4rem}
  .kpi{flex:1 1 calc(33.333% - .4rem);max-width:none;padding:.7rem .4rem;min-width:0}
  .kpi:last-child:nth-child(3n+1){flex-basis:100%}
  .kpi:nth-last-child(2):nth-child(3n+1),
  .kpi:last-child:nth-child(3n+2){flex-basis:calc(50% - .4rem)}
  .kpi-value{font-size:clamp(.72rem, 4.1vw, 1.15rem);overflow-wrap:anywhere}
  .kpi-label{font-size:clamp(.42rem, 2.2vw, .58rem);letter-spacing:.03em;line-height:1.2;margin-bottom:4px;overflow-wrap:break-word;hyphens:auto}
  .kpi-sub{font-size:clamp(.4rem, 2vw, .58rem);line-height:1.2;margin-top:3px;overflow-wrap:break-word;hyphens:auto}
  .top-kpi-row{grid-template-columns:repeat(3,1fr) !important;gap:.4rem}
  .top-kpi:last-child:nth-child(3n+1){grid-column:1 / -1}
  .top-kpi:nth-last-child(2):nth-child(3n+1){grid-column:span 2}
  .top-kpi{padding:.65rem .4rem;min-width:0}
  .top-kpi .kpi-value{font-size:clamp(.68rem, 3.4vw, .95rem);overflow-wrap:anywhere}
  .top-kpi .kpi-label{font-size:clamp(.4rem, 2.1vw, .55rem);line-height:1.2}
  .top-kpi .kpi-value.kv-txt{font-size:clamp(.6rem, 2.9vw, .85rem) !important}
  .filter-bar{padding:.5rem .75rem;gap:.4rem}
  .suc-label{font-size:.6rem;margin-right:0}
  .suc-hint{display:none}
  .suc-btns{gap:4px}
  .suc-btn,.mes-btn{padding:3px 8px;font-size:.6rem}
  .btn-all{padding:4px 10px;font-size:.6rem}
  .suc-sep{height:14px}
  .tab-nav-btn{padding:.7rem .9rem;font-size:.74rem}
  header{column-gap:.5rem;padding:1rem 1rem}
  .header-logo img{height:4.5rem}
}
</style>
</head>
<body>
<div class="sticky-header">
<header>
  <div class="header-left">
    <h1>Skin Love — Análisis de Ventas</h1>
    <p><strong>Período __MES_HEADER__</strong></p>
    <p>Elaborado con información al día __FECHA_INFO__</p>
  </div>
  <div class="header-logo"><img src="data:image/png;base64,__LOGO_B64__" alt="Skin Love"></div>
  <div class="header-right">
    <span class="hdate">Reporte Generado: __FECHA_VALOR__</span>
    <div class="hbadge">DASHBOARD v1.0</div>
  </div>
</header>
<div class="filter-bar" id="filter-bar-suc">
  <span class="suc-label">Canal de Venta</span>
  <div class="suc-btns" id="suc-btns"></div>
  <div class="suc-sep"></div>
  <button class="btn-all" onclick="toggleAll()">Todos / Ninguno</button>
  <span class="suc-hint">Clic: selección única · Ctrl/⌘ + clic: selección múltiple</span>
</div>
<div class="filter-bar" id="filter-bar-mes">
  <span class="suc-label" style="color:#3D3D3D">Meses · Histórico</span>
  <div class="suc-btns" id="mes-btns"></div>
  <div class="suc-sep"></div>
  <button class="btn-all dark" onclick="toggleAllMeses()">Todos / Ninguno</button>
</div>
<div class="tabs-nav">
  <button class="tab-nav-btn active" id="tabnav-pronostico"    onclick="switchTab('pronostico')">Pronóstico Mes en Curso</button>__TABS_NAV_EXTRA__
  <button class="tab-nav-btn"        id="tabnav-resumenactual"    onclick="switchTab('resumenactual')">Resumen Mes Actual</button>
  <button class="tab-nav-btn"        id="tabnav-resumenacumulado" onclick="switchTab('resumenacumulado')">Resumen Histórico</button>
  <button class="tab-nav-btn"        id="tabnav-toparticulos" onclick="switchTab('toparticulos')">Top 10 Artículos</button>
  <button class="tab-nav-btn"        id="tabnav-lineascategoria" onclick="switchTab('lineascategoria')">Líneas y Categorías</button>
  <button class="tab-nav-btn"        id="tabnav-fabricantes" onclick="switchTab('fabricantes')">Fabricantes</button>
</div>
</div>
<div id="tab-pronostico" class="tab-content active">
<div class="main">
  <div class="tc" style="margin-bottom:.9rem">
    <div class="card-head">
      <div><div class="card-title">Pronóstico de Alcance</div><div class="card-sub">Ventas estimadas al cierre de mes, por canal de venta</div></div>
      <span class="note-bol">Proyección del ritmo de venta observado en el mes</span>
    </div>
    <div class="top-kpi-row" style="grid-template-columns:repeat(4,1fr)">
      <div class="top-kpi" style="--ac:#6B6B6B"><div class="kpi-label">Unidades</div><div class="kpi-value" id="p-unidades">—</div><div class="kpi-sub">Acumulado mes en curso</div></div>
      <div class="top-kpi" style="--ac:#3D3D3D"><div class="kpi-label">Ventas $</div><div class="kpi-value" id="p-ventas">—</div><div class="kpi-sub">Acumulado mes en curso</div></div>
      <div class="top-kpi" style="--ac:#5A5A5A"><div class="kpi-label">Utilidad</div><div class="kpi-value brand" id="p-util">—</div><div class="kpi-sub" id="p-util-sub">Margen: —</div></div>
      <div class="top-kpi" style="--ac:#aa7300"><div class="kpi-label">Pronóstico de cierre</div><div class="kpi-value gold" id="p-pronostico">—</div><div class="kpi-sub" id="p-pronostico-sub">Suma de pronósticos</div></div>
    </div>
    <div class="table-scale-wrap">
      <table>
        <thead><tr><th>Canal de Venta</th><th class="r">Unidades</th><th class="r">Venta $</th><th class="r">Utilidad</th><th class="r">Margen</th><th class="r">Pronóstico de cierre</th></tr></thead>
        <tbody id="tabla-pronostico"></tbody>
      </table>
    </div>
  </div>
  <div class="hist-card">
    <div class="card-head">
      <div><div class="card-title">Ventas acumuladas y Pronóstico de cierre</div><div class="card-sub">Barras: venta real acumulada · Línea: pronóstico al cierre del mes</div></div>
    </div>
    <div class="cw" id="pron-chart-wrap" style="height:420px"><canvas id="chart-pronostico"></canvas></div>
  </div>
</div>
</div>
<div id="tab-resumenactual" class="tab-content">
<div class="main">
  <p class="kpi-note">* Esta pestaña sólo muestra información correspondiente al mes en curso (__MES_HEADER__).</p>
  <div class="kpi-grid">
    <div class="kpi" style="--ac:#9A9A9A"><div class="kpi-label">Unidades vendidas</div><div class="kpi-value" id="k-uni">—</div><div class="kpi-sub">Ventas mes en curso</div></div>
    <div class="kpi" style="--ac:#808080"><div class="kpi-label">Ventas $</div><div class="kpi-value brand" id="k-ventas">—</div><div class="kpi-sub" id="k-ventas-s">—</div></div>
    <div class="kpi" style="--ac:#5A5A5A"><div class="kpi-label">Utilidad</div><div class="kpi-value brand" id="k-util">—</div><div class="kpi-sub" id="k-util-sub">Margen: —</div></div>
    <div class="kpi" style="--ac:#3D3D3D"><div class="kpi-label">Tickets</div><div class="kpi-value" id="k-tkt">—</div><div class="kpi-sub" id="k-tkt-sub">Volumen de ventas</div></div>
    <div class="kpi" style="--ac:#A8A8A8"><div class="kpi-label">Ventas $ promedio por ticket</div><div class="kpi-value" id="k-vtkt">—</div><div class="kpi-sub" id="k-vtkt-sub">Ventas $ ÷ Tickets</div></div>
    <div class="kpi" style="--ac:#C4C4C4"><div class="kpi-label">Unidades promedio por Ticket</div><div class="kpi-value" id="k-utkt">—</div><div class="kpi-sub" id="k-utkt-sub">Unidades ÷ Tickets</div></div>
  </div>
  <div class="tc">
    <div class="card-head">
      <div><div class="card-title">Resumen de Ventas · Mes en curso</div><div class="card-sub">Detalle diario según canales de venta seleccionados</div></div>
      <span class="note-bol">Desempeño de ventas por día</span>
    </div>
    <div class="table-scale-wrap">
    <table>
      <thead><tr><th>Fecha</th><th>Día</th><th class="r">Unidades</th><th class="r">Ventas $</th><th class="r">Utilidad</th><th class="r">Margen</th><th class="r">Tickets</th></tr></thead>
      <tbody id="tabla-body"></tbody>
    </table>
    </div>
  </div>
  <div class="charts-row">
    <div class="cc"><div class="card-head" style="margin-bottom:.4rem"><div><div class="card-title">Ventas $</div><div class="card-sub">Volumen diario · Canales de venta seleccionados</div></div></div><div class="cw"><canvas id="chart-ventas"></canvas></div></div>
    <div class="cc"><div class="card-head" style="margin-bottom:.4rem"><div><div class="card-title">No. de Tickets</div><div class="card-sub">Volumen diario · Canales de venta seleccionados</div></div></div><div class="cw"><canvas id="chart-tickets"></canvas></div></div>
  </div>
  <hr class="section-divider">
  <div class="tc" style="margin-bottom:.9rem">
    <div class="card-head">
      <div><div class="card-title">Comparativo por Canal de Venta · Mes en Curso</div><div class="card-sub">Unidades, ventas, utilidad y margen acumulados del mes</div></div>
      <span class="note-bol">Sólo __MES_HEADER__</span>
    </div>
    <div class="table-scale-wrap">
    <table>
      <thead><tr><th>Canal de Venta</th><th class="r">Unidades</th><th class="r">Ventas $</th><th class="r">Utilidad</th><th class="r">Margen</th><th class="r">Tickets</th></tr></thead>
      <tbody id="tabla-sucursal-actual"></tbody>
    </table>
    </div>
  </div>
  <div class="hist-card">
    <div class="card-head">
      <div><div class="card-title">Comparativo por Canal de Venta</div><div class="card-sub">Métrica seleccionada + Margen % · Mes en curso</div></div>
      <div class="metric-tabs">
        <button class="tab-btn active" id="btn-hs-ventas"   onclick="changeSucMetric('ventas')">Ventas c/Desc</button>
        <button class="tab-btn"        id="btn-hs-utilidad" onclick="changeSucMetric('utilidad')">Utilidad</button>
        <button class="tab-btn"        id="btn-hs-unidades" onclick="changeSucMetric('unidades')">Unidades</button>
        <button class="tab-btn"        id="btn-hs-tickets"  onclick="changeSucMetric('tickets')">Tickets</button>
      </div>
    </div>
    <div class="cw" id="sucursalactual-chart-wrap" style="height:340px"><canvas id="chart-sucursalactual"></canvas></div>
  </div>
  <hr class="section-divider">
  <div class="lines-container">
    <div class="lines-table-box">
      <div class="card-head">
        <div><div class="card-title">Resumen de Ventas por Línea · Mes en Curso</div><div class="card-sub">Sólo __MES_HEADER__</div></div>
      </div>
      <div class="table-scale-wrap">
        <table>
          <thead><tr><th>Línea</th><th class="r">Unidades</th><th class="r">Ventas $</th><th class="r">Utilidad</th><th class="r">Margen</th></tr></thead>
          <tbody id="tabla-lineas-actual"></tbody>
        </table>
      </div>
    </div>
    <div class="lines-chart-box">
      <div class="card-head" style="margin-bottom:.7rem">
        <div><div class="card-title">Porcentaje de Participación (%)</div><div class="card-sub">Mes en curso</div></div>
        <div class="metric-tabs">
          <button class="tab-btn active" id="btn-ma-ventas"   onclick="changeLineMetricActual('ventas')">Ventas c/Desc</button>
          <button class="tab-btn"        id="btn-ma-unidades" onclick="changeLineMetricActual('unidades')">Unidades</button>
          <button class="tab-btn"        id="btn-ma-utilidad" onclick="changeLineMetricActual('utilidad')">Utilidad</button>
        </div>
      </div>
      <div class="cw bars-horizontal" id="lineasactual-chart-wrap"><canvas id="chart-lineasactual"></canvas></div>
    </div>
  </div>
</div>
</div>
<div id="tab-resumenacumulado" class="tab-content">
<div class="main">
  <p class="kpi-note">* Las tarjetas muestran información de acuerdo a lo seleccionado en los segmentadores.</p>
  <div class="kpi-grid">
    <div class="kpi" style="--ac:#9A9A9A"><div class="kpi-label">Unidades vendidas</div><div class="kpi-value" id="ka-uni">—</div><div class="kpi-sub">Acumulado según selección</div></div>
    <div class="kpi" style="--ac:#808080"><div class="kpi-label">Ventas $</div><div class="kpi-value brand" id="ka-ventas">—</div><div class="kpi-sub" id="ka-ventas-s">—</div></div>
    <div class="kpi" style="--ac:#5A5A5A"><div class="kpi-label">Utilidad</div><div class="kpi-value brand" id="ka-util">—</div><div class="kpi-sub" id="ka-util-sub">Margen: —</div></div>
    <div class="kpi" style="--ac:#3D3D3D"><div class="kpi-label">Tickets</div><div class="kpi-value" id="ka-tkt">—</div><div class="kpi-sub" id="ka-tkt-sub">Volumen de ventas</div></div>
    <div class="kpi" style="--ac:#A8A8A8"><div class="kpi-label">Ventas $ promedio por ticket</div><div class="kpi-value" id="ka-vtkt">—</div><div class="kpi-sub" id="ka-vtkt-sub">Ventas $ ÷ Tickets</div></div>
    <div class="kpi" style="--ac:#C4C4C4"><div class="kpi-label">Unidades promedio por Ticket</div><div class="kpi-value" id="ka-utkt">—</div><div class="kpi-sub" id="ka-utkt-sub">Unidades ÷ Tickets</div></div>
  </div>
  <p class="kpi-note" id="nota-consolidado" style="display:none"></p>
  <div class="tc">
    <div class="card-head">
      <div><div class="card-title">Ventas Históricas Mensuales</div><div class="card-sub">Acumulado por mes · Canales de venta y meses seleccionados</div></div>
    </div>
    <div class="table-scale-wrap">
    <table>
      <thead><tr><th>Período</th><th class="r">Unidades</th><th class="r">Ventas $</th><th class="r">Utilidad</th><th class="r">Margen</th><th class="r">Tickets</th></tr></thead>
      <tbody id="tabla-historico"></tbody>
    </table>
    </div>
  </div>
  <div class="hist-card">
    <div class="card-head">
      <div><div class="card-title">Tendencia Histórica Mensual</div><div class="card-sub">Ventas mensuales y Tasa de incremento vs. mes anterior · Canales de venta y meses seleccionados</div></div>
      <div class="metric-tabs">
        <button class="tab-btn active" id="btn-h-ventas"   onclick="changeHistMetric('ventas')">Ventas c/Desc</button>
        <button class="tab-btn"        id="btn-h-utilidad" onclick="changeHistMetric('utilidad')">Utilidad</button>
        <button class="tab-btn"        id="btn-h-unidades" onclick="changeHistMetric('unidades')">Unidades</button>
        <button class="tab-btn"        id="btn-h-tickets"  onclick="changeHistMetric('tickets')">Tickets</button>
      </div>
    </div>
    <div class="cw hist-chart"><canvas id="chart-historico"></canvas></div>
    <p class="kpi-note" id="nota-tendencia" style="display:none;margin:.55rem 0 0"></p>
  </div>
  <div class="charts-row" style="margin-top:.9rem">
    <div class="cc">
      <div class="card-head" style="margin-bottom:.4rem">
        <div><div class="card-title" id="histcomp-anio-title">Ventas: Mes en Curso vs. Mismo Mes Año Anterior</div><div class="card-sub">Comparativo interanual · Sigue el mes marcado en el segmentador (si hay uno solo) · Canales de venta seleccionados</div></div>
        <span class="pill" id="histcomp-anio-var">—</span>
      </div>
      <div class="cw" id="histcomp-anio-wrap"><canvas id="chart-hist-comp-anio"></canvas></div>
    </div>
    <div class="cc">
      <div class="card-head" style="margin-bottom:.4rem">
        <div><div class="card-title" id="histcomp-mes-title">Ventas: Mes en Curso vs. Mes Anterior</div><div class="card-sub">Comparativo mensual · Sigue el mes marcado en el segmentador (si hay uno solo) · Canales de venta seleccionados</div></div>
        <span class="pill" id="histcomp-mes-var">—</span>
      </div>
      <div class="cw" id="histcomp-mes-wrap"><canvas id="chart-hist-comp-mes"></canvas></div>
    </div>
  </div>
  <hr class="section-divider">
  <div class="lines-container">
    <div class="lines-table-box">
      <div class="card-head">
        <div><div class="card-title">Resumen de Ventas Acumuladas por Línea</div><div class="card-sub">Acumulado de ventas según selección</div></div>
      </div>
      <div class="table-scale-wrap">
        <table>
          <thead><tr><th>Línea</th><th class="r">Unidades</th><th class="r">Ventas $</th><th class="r">Utilidad</th><th class="r">Margen</th></tr></thead>
          <tbody id="tabla-lineas"></tbody>
        </table>
      </div>
    </div>
    <div class="lines-chart-box">
      <div class="card-head" style="margin-bottom:.7rem">
        <div><div class="card-title">Porcentaje de Participación (%)</div><div class="card-sub">Proporción porcentual de acuerdo a Ventas, Unidades o Utilidad</div></div>
        <div class="metric-tabs">
          <button class="tab-btn active" id="btn-m-ventas"   onclick="changeLineMetric('ventas')">Ventas c/Desc</button>
          <button class="tab-btn"        id="btn-m-unidades" onclick="changeLineMetric('unidades')">Unidades</button>
          <button class="tab-btn"        id="btn-m-utilidad" onclick="changeLineMetric('utilidad')">Utilidad</button>
        </div>
      </div>
      <div class="cw bars-horizontal" id="lineas-chart-wrap"><canvas id="chart-lineas"></canvas></div>
    </div>
  </div>
</div>
</div>
<div id="tab-toparticulos" class="tab-content">
<div class="main">
  <div class="top-kpi-row">
    <div class="top-kpi" style="--ac:#6B6B6B"><div class="kpi-label">Unidades vendidas Top 10</div><div class="kpi-value" id="tk-uni">—</div><div class="kpi-sub">Suma de top 10</div></div>
    <div class="top-kpi" style="--ac:#aa7300"><div class="kpi-label">Ventas $ Top 10</div><div class="kpi-value gold" id="tk-ventas">—</div><div class="kpi-sub">Suma de top 10</div></div>
    <div class="top-kpi" style="--ac:#1A7A4A"><div class="kpi-label">% participación sobre ventas $</div><div class="kpi-value" id="tk-margen" style="color:#1A7A4A">—</div><div class="kpi-sub">Ventas $ Top 10 ÷ Ventas $ acumuladas</div></div>
  </div>
  <div class="top-layout">
    <div class="top-table-box">
      <div class="card-head">
        <div><div class="card-title">Ranking de Artículos</div><div class="card-sub">Top 10 por Unidades Vendidas</div></div>
        <span class="note-bol">Acorde a selección</span>
      </div>
      <div class="table-scale-wrap">
      <table>
        <thead>
          <tr>
            <th style="width:36px;text-align:center">#</th>
            <th>Artículo</th>
            <th>Descripción</th>
            <th>Fabricante</th>
            <th class="r">Unidades</th>
            <th class="r">Venta $</th>
            <th class="r">% Venta</th>
          </tr>
        </thead>
        <tbody id="tabla-top"></tbody>
      </table>
      </div>
    </div>
    <div class="top-bubble-box">
      <div class="card-head" style="margin-bottom:.5rem">
        <div>
          <div class="card-title">Mapa de Desempeño · Top 10</div>
          <div class="card-sub">Eje X: Ventas $ · Eje Y: Margen (%) · Tamaño burbuja: Unidades vendidas</div>
        </div>
      </div>
      <div class="cw bubble-chart"><canvas id="chart-bubble"></canvas></div>
    </div>
  </div>
</div>
</div>
<div id="tab-lineascategoria" class="tab-content">
<div class="main">
  <div class="top-kpi-row">
    <div class="top-kpi" style="--ac:#808080"><div class="kpi-label">Línea Líder</div><div class="kpi-value brand kv-txt" id="tk2-linea-nombre" style="font-size:1.05rem">—</div><div class="kpi-sub" id="tk2-linea-sub">—</div></div>
    <div class="top-kpi" style="--ac:#6B6B6B"><div class="kpi-label">Categoría Líder</div><div class="kpi-value kv-txt" id="tk2-cat-nombre" style="color:#6B6B6B;font-size:1.05rem">—</div><div class="kpi-sub" id="tk2-cat-sub">—</div></div>
    <div class="top-kpi" style="--ac:#1A7A4A"><div class="kpi-label">Margen · Categoría Líder</div><div class="kpi-value" id="tk2-cat-margen" style="color:#1A7A4A">—</div><div class="kpi-sub">Utilidad ÷ Ventas de la categoría líder</div></div>
  </div>
  <div class="top-table-box">
    <div class="card-head">
      <div><div class="card-title">Detalle por Línea y Categoría</div><div class="card-sub">Clic en una línea para desplegar sus categorías a detalle</div></div>
      <span class="note-bol">Acorde a canales de venta y meses seleccionados</span>
    </div>
    <div class="table-scale-wrap">
    <table>
      <thead>
        <tr>
          <th style="width:26px"></th>
          <th>Línea / Categoría</th>
          <th class="r">Unidades</th>
          <th class="r">Ventas $</th>
          <th class="r">Utilidad</th>
          <th class="r">Margen</th>
        </tr>
      </thead>
      <tbody id="tabla-lineascategoria"></tbody>
    </table>
    </div>
  </div>
  <div class="hist-card" style="margin-top:1rem">
    <div class="card-head">
      <div><div class="card-title">Mapa de Línea &gt; Categoría</div><div class="card-sub">Top 3 categorías por línea </div></div>
      <span class="legend-gradient">Menor venta<span class="legend-bar"></span>Mayor venta</span>
    </div>
    <div class="cw" id="lc-chart-wrap" style="height:360px"><canvas id="chart-treemap"></canvas></div>
  </div>
</div>
</div>
<div id="tab-fabricantes" class="tab-content">
<div class="main">
  <div class="top-kpi-row">
    <div class="top-kpi" style="--ac:#808080"><div class="kpi-label">Ventas $ · Fabricantes al 50%</div><div class="kpi-value brand" id="tk3-ventas">—</div><div class="kpi-sub">Suma de fabricantes hasta cubrir el 50%</div></div>
    <div class="top-kpi" style="--ac:#1A7A4A"><div class="kpi-label">Margen · Fabricantes al 50%</div><div class="kpi-value" id="tk3-margen" style="color:#1A7A4A">—</div><div class="kpi-sub">Utilidad ÷ Ventas del grupo</div></div>
    <div class="top-kpi" style="--ac:#aa7300"><div class="kpi-label">Fabricantes incluidos</div><div class="kpi-value gold" id="tk3-pct">—</div><div class="kpi-sub" id="tk3-pct-sub">Necesarios para alcanzar el 50% de ventas</div></div>
  </div>
  <div class="hist-card" style="margin-bottom:1rem">
    <div class="card-head">
      <div><div class="card-title">Fabricantes que concentran el 50% de las Ventas</div><div class="card-sub">% de participación en Ventas $</div></div>
    </div>
    <div class="cw" id="fab-chart-wrap" style="height:420px"><canvas id="chart-fabricantes"></canvas></div>
  </div>
  <div class="top-table-box">
    <div class="card-head">
      <div><div class="card-title">Detalle Completo por Fabricante</div><div class="card-sub">Todos los fabricantes de la selección, agrupados en cuartiles según su venta acumulada · clic en un cuartil para desplegar el detalle</div></div>
      <span class="note-bol">Acorde a canales de venta y meses seleccionados</span>
    </div>
    <div class="table-scale-wrap">
    <table>
      <thead>
        <tr>
          <th style="width:26px"></th>
          <th>Fabricante</th>
          <th class="r">Ventas $</th>
          <th class="r">Margen</th>
          <th class="r">% Participación</th>
        </tr>
      </thead>
      <tbody id="tabla-fabricantes"></tbody>
    </table>
    </div>
  </div>
</div>
</div>
__TABS_CONTENT_EXTRA__
<footer style="text-align:center;padding:1.4rem 2rem 2rem;font-size:.7rem;color:#a3a3a3;border-top:1px solid #ebebeb;margin-top:1rem">
  Elaborado por el Equipo de Planeación y Análisis de la información.
</footer>
<script>
document.addEventListener("DOMContentLoaded", function() {
    const RAW         = __DATA_JSON__;
    const LINEAS      = __LINEA_JSON__;
    const HISTORICO   = __HISTORICO_JSON__;
    const TOP_ART     = __TOP_ART_JSON__;
    const LINEAS_CAT  = __LINEAS_CAT_JSON__;
    const FABRICANTES = __FABRICANTES_JSON__;
    const PRONOSTICO  = __PRONOSTICO_JSON__;
    const SUCS        = __SUCURSALES_JSON__;
    const PERIODOS    = __PERIODOS_JSON__;
    const RESUMENES_ANT = __RESUMENES_ANT_JSON__;
    const CURRENT_PERIOD = __CURRENT_PERIOD_JSON__;
    const TICKETS_CONS = __TICKETS_CONS_JSON__;
    const HAY_CONSOLIDADO = Object.keys(TICKETS_CONS).length > 0;
    const NOTA_CONSOLIDADO = '* Los períodos anteriores a Sep-2026 no cuentan con detalle de tickets por canal de venta, por esta razón las Ventas y Unidades Promedio por Ticket se calculan, para estos períodos, usando los totales.';
    const NOTA_TENDENCIA = '* Los períodos sin desglose se muestran como una barra única con el total acumulado. Al filtrar canales el cálculo de la tasa de crecimiento se omite con el salto entre períodos sin desglose y períodos con desglose.';
    function periodosConsolidadosEn(data){
        return [...new Set(data.map(r => r.PeriodoLabel))]
            .filter(p => TICKETS_CONS[p] !== undefined).sort();
    }
    const SUCURSALES_POR_PERIODO = {};
    const PERIODOS_POR_SUCURSAL  = {};
    HISTORICO.forEach(r => {
        if(!(r.ventas !== 0 || r.unidades !== 0 || r.tickets > 0)) return;
        if(!SUCURSALES_POR_PERIODO[r.PeriodoLabel]) SUCURSALES_POR_PERIODO[r.PeriodoLabel] = new Set();
        SUCURSALES_POR_PERIODO[r.PeriodoLabel].add(r.NombreSucursal);
        if(!PERIODOS_POR_SUCURSAL[r.NombreSucursal]) PERIODOS_POR_SUCURSAL[r.NombreSucursal] = new Set();
        PERIODOS_POR_SUCURSAL[r.NombreSucursal].add(r.PeriodoLabel);
    });
    const RAW_SUCS = new Set(RAW.filter(r => r.ventas !== 0 || r.unidades !== 0 || r.tickets > 0).map(r => r.NombreSucursal));
    const MESES_ABR = ['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic'];
    const PALETTE = [
        '#4D4D4D','#2196A8','#E05C2A','#3A7D44','#C4922A',
        '#1A559E','#B03060','#4E8C6E','#6B4226','#5C5FA8',
        '#8E3A59','#2E7D60','#A04010','#3D6B99','#7A6F1E',
        '#9C3D6B','#1E6B7A','#5A3A8E','#3D7A4A','#8E5A1E',
    ];
    const PALETTE_LINEAS = [
        '#2B2B2B','#3D3D3D','#5A5A5A','#6B6B6B','#808080',
        '#9A9A9A','#A8A8A8','#B8B8B8','#C4C4C4','#D4D4D4'
    ];
    const PALETTE_BUBBLE = [
        '#4D4D4D','#2196A8','#E05C2A','#3A7D44','#C4922A',
        '#1A559E','#B03060','#4E8C6E','#6B4226','#5C5FA8'
    ];
    function buildColorMap(list, palette) {
        const map = {};
        list.forEach((name, i) => { map[name] = palette[i % palette.length]; });
        return map;
    }
    const ALL_LINEAS = [...new Set(LINEAS.map(r => r.Línea))].sort();
    const LC = buildColorMap(ALL_LINEAS, PALETTE_LINEAS);
    let SC = {};
    const DAYS = ['Dom','Lun','Mar','Mié','Jue','Vie','Sáb'];
    let active      = new Set(SUCS);
    let activeWanted = new Set(SUCS);
    let activeMeses = new Set(PERIODOS);
    let sucursalExclusiva = false;
    let mesExclusivo = false;
    let charts      = {};
    let currentLineMetric = 'ventas';
    let currentLineMetricActual = 'ventas';
    let currentHistMetric = 'ventas';
    let currentSucMetric  = 'ventas';
    let currentTab = 'pronostico';
    const fM  = v => '$'+(v>=1e6?(v/1e6).toFixed(1)+'M':v>=1e3?(v/1e3).toFixed(0)+'K':Math.round(v));
    const fF  = v => '$'+Math.round(v).toLocaleString('es-MX');
    const fP  = v => (v*100).toFixed(2)+'%';
    const fN  = v => Math.round(v).toLocaleString('es-MX');
    const PREMIUM_TOOLTIP_OPTS = {
        backgroundColor:'#ffffff',titleColor:'#212121',bodyColor:'#4d4d4d',
        borderColor:'#e6e6e6',borderWidth:1,padding:10,cornerRadius:8,
        boxPadding:6,usePointStyle:true,
        titleFont:{family:"'Segoe UI', sans-serif",weight:'bold',size:12},
        bodyFont:{family:"'Segoe UI', sans-serif",size:12}
    };
    Chart.register(ChartDataLabels);
    function fitTables(){
        document.querySelectorAll('.table-scale-wrap').forEach(wrap => {
            const table = wrap.querySelector('table');
            if(!table) return;
            table.style.transform = 'none';
            wrap.style.height = 'auto';
            const availWidth  = wrap.clientWidth;
            const neededWidth = table.scrollWidth;
            if(availWidth > 0 && neededWidth > availWidth){
                const scale = availWidth / neededWidth;
                table.style.transform = `scale(${scale})`;
                wrap.style.height = (table.offsetHeight * scale) + 'px';
            }
        });
    }
    window.addEventListener('resize', () => {
        clearTimeout(window._fitTablesTimer);
        window._fitTablesTimer = setTimeout(fitTables, 150);
    });
    window.switchTab = function(tab) {
        currentTab = tab;
        document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));
        document.querySelectorAll('.tab-nav-btn').forEach(el => el.classList.remove('active'));
        document.getElementById('tab-' + tab).classList.add('active');
        document.getElementById('tabnav-' + tab).classList.add('active');
        const usaMeses = (tab !== 'pronostico' && tab !== 'resumenactual');
        document.getElementById('filter-bar-mes').style.display = usaMeses ? 'flex' : 'none';
        const usaSucursales = (tab !== 'pronostico');
        document.getElementById('filter-bar-suc').style.display = usaSucursales ? 'flex' : 'none';
        if(usaSucursales) refreshSucursalesDisponibilidad(false, true);
        if (tab === 'resumenactual') { update(); updateResumenActualExtra(); }
        if (tab === 'resumenacumulado') updateAcumulado();
        if (tab === 'pronostico') updatePronostico();
        if (tab === 'toparticulos') updateTopArticulos();
        if (tab === 'lineascategoria') updateLineasCategoria();
        if (tab === 'fabricantes') updateFabricantes();
        if (tab.startsWith('resumen_')) updateResumenAnterior(tab.replace('resumen_',''));
        setTimeout(fitTables, 50);
    };
    function buildButtons(){
        SC = buildColorMap(SUCS, PALETTE);
        const wrap = document.getElementById('suc-btns');
        wrap.innerHTML = '';
        SUCS.forEach(s => {
            const btn = document.createElement('button');
            btn.className = 'suc-btn active';
            btn.dataset.sucursal = s;
            btn.innerHTML = `<span style="display:inline-block;width:7px;height:7px;border-radius:50%;background:${SC[s]};margin-right:6px;flex-shrink:0"></span>${s}`;
            btn.addEventListener('click', (ev) => {
                if(btn.disabled) return;
                const multi = ev.ctrlKey || ev.metaKey || ev.shiftKey;
                if(multi){
                    if(active.has(s)){ active.delete(s); activeWanted.delete(s); btn.classList.remove('active'); }
                    else { active.add(s); activeWanted.add(s); btn.classList.add('active'); }
                    sucursalExclusiva = false;
                    refreshMesesDisponibilidad(false);
                } else {
                    active.clear();
                    activeWanted.clear();
                    active.add(s);
                    activeWanted.add(s);
                    document.querySelectorAll('#suc-btns .suc-btn').forEach(b => {
                        b.classList.toggle('active', b.dataset.sucursal === s);
                    });
                    sucursalExclusiva = true;
                    refreshMesesDisponibilidad(false);
                    const debeExpandirMeses = !mesExclusivo || activeMeses.size === 0;
                    if(debeExpandirMeses){
                        refreshMesesDisponibilidad(true);
                        mesExclusivo = false;
                    }
                }
                updateAll();
            });
            wrap.appendChild(btn);
        });
    }
    function buildMesButtons(){
        const wrap = document.getElementById('mes-btns');
        wrap.innerHTML = '';
        PERIODOS.forEach(p => {
            const btn = document.createElement('button');
            btn.className = 'mes-btn active';
            btn.textContent = p;
            btn.dataset.periodo = p;
            btn.addEventListener('click', (ev) => {
                if(btn.disabled) return;
                const multi = ev.ctrlKey || ev.metaKey || ev.shiftKey;
                if(multi){
                    if(activeMeses.has(p)){ activeMeses.delete(p); btn.classList.remove('active'); }
                    else { activeMeses.add(p); btn.classList.add('active'); }
                    mesExclusivo = false;
                    refreshSucursalesDisponibilidad(false);
                } else {
                    activeMeses.clear();
                    activeMeses.add(p);
                    document.querySelectorAll('#mes-btns .mes-btn').forEach(b => {
                        b.classList.toggle('active', b.dataset.periodo === p);
                    });
                    mesExclusivo = true;
                    refreshSucursalesDisponibilidad(false);
                    const debeExpandirSucursales = !sucursalExclusiva || active.size === 0;
                    if(debeExpandirSucursales){
                        refreshSucursalesDisponibilidad(true);
                        sucursalExclusiva = false;
                    }
                }
                if(currentTab === 'resumenacumulado') updateAcumulado();
                if(currentTab === 'toparticulos') updateTopArticulos();
                if(currentTab === 'lineascategoria') updateLineasCategoria();
                if(currentTab === 'fabricantes') updateFabricantes();
            });
            wrap.appendChild(btn);
        });
    }
    function refreshSucursalesDisponibilidad(autoSeleccionar, restaurarDesdeMemoria){
        let cambio = false;
        const criterioMesActual = (currentTab === 'resumenactual');
        document.querySelectorAll('#suc-btns .suc-btn').forEach(btn => {
            const s = btn.dataset.sucursal;
            let esValida;
            if(criterioMesActual){
                esValida = RAW_SUCS.has(s);
            } else {
                esValida = activeMeses.size === 0;
                if(!esValida){
                    for(const p of activeMeses){
                        if((SUCURSALES_POR_PERIODO[p] || new Set()).has(s)){ esValida = true; break; }
                    }
                }
            }
            btn.disabled = !esValida;
            btn.classList.toggle('oculto', !esValida);
            if(!esValida && active.has(s)){
                active.delete(s);
                btn.classList.remove('active');
                cambio = true;
            } else if(esValida && !active.has(s)){
                const debeActivar = autoSeleccionar || (restaurarDesdeMemoria && activeWanted.has(s));
                if(debeActivar){
                    active.add(s);
                    activeWanted.add(s);
                    btn.classList.add('active');
                    cambio = true;
                }
            }
        });
        return cambio;
    }
    function refreshMesesDisponibilidad(autoSeleccionar){
        let cambio = false;
        document.querySelectorAll('#mes-btns .mes-btn').forEach(btn => {
            const p = btn.dataset.periodo;
            let esValido = active.size === 0;
            if(!esValido){
                for(const s of active){
                    if((PERIODOS_POR_SUCURSAL[s] || new Set()).has(p)){ esValido = true; break; }
                }
            }
            btn.disabled = !esValido;
            btn.classList.toggle('oculto', !esValido);
            if(!esValido && activeMeses.has(p)){
                activeMeses.delete(p);
                btn.classList.remove('active');
                cambio = true;
            } else if(esValido && autoSeleccionar && !activeMeses.has(p)){
                activeMeses.add(p);
                btn.classList.add('active');
                cambio = true;
            }
        });
        return cambio;
    }
    window.toggleAll = function(){
        const btns = [...document.querySelectorAll('#suc-btns .suc-btn')];
        const seleccionables = btns.filter(b => !b.disabled);
        const todasActivas = seleccionables.length>0 && seleccionables.every(b => b.classList.contains('active'));
        if(todasActivas){
            seleccionables.forEach(b => { active.delete(b.dataset.sucursal); activeWanted.delete(b.dataset.sucursal); b.classList.remove('active'); });
        } else {
            seleccionables.forEach(b => { active.add(b.dataset.sucursal); activeWanted.add(b.dataset.sucursal); b.classList.add('active'); });
        }
        sucursalExclusiva = false;
        refreshMesesDisponibilidad();
        updateAll();
    };
    window.toggleAllMeses = function(){
        const btns = [...document.querySelectorAll('#mes-btns .mes-btn')];
        const seleccionables = btns.filter(b => !b.disabled);
        const todosActivos = seleccionables.length>0 && seleccionables.every(b => b.classList.contains('active'));
        if(todosActivos){
            seleccionables.forEach(b => { activeMeses.delete(b.dataset.periodo); b.classList.remove('active'); });
        } else {
            seleccionables.forEach(b => { activeMeses.add(b.dataset.periodo); b.classList.add('active'); });
        }
        mesExclusivo = false;
        refreshSucursalesDisponibilidad();
        if(currentTab === 'resumenacumulado') updateAcumulado();
        if(currentTab === 'toparticulos') updateTopArticulos();
        if(currentTab === 'lineascategoria') updateLineasCategoria();
        if(currentTab === 'fabricantes') updateFabricantes();
    };
    function updateAll(){
        if(currentTab === 'resumenactual') { update(); updateResumenActualExtra(); }
        if(currentTab === 'resumenacumulado') updateAcumulado();
        if(currentTab === 'toparticulos') updateTopArticulos();
        if(currentTab === 'lineascategoria') updateLineasCategoria();
        if(currentTab === 'fabricantes') updateFabricantes();
        if(currentTab && currentTab.startsWith('resumen_')) updateResumenAnterior(currentTab.replace('resumen_',''));
    }
    function getFiltered()          { return RAW.filter(r => active.has(r.NombreSucursal)); }
    function getFilteredLineas()    {
        return LINEAS.filter(r =>
            active.has(r.NombreSucursal) && activeMeses.has(r.PeriodoLabel)
        );
    }
    function getFilteredLineasActual() {
        return LINEAS.filter(r =>
            active.has(r.NombreSucursal) && r.PeriodoLabel === CURRENT_PERIOD
        );
    }
    function getFilteredHist()      {
        return HISTORICO.filter(r =>
            active.has(r.NombreSucursal) && activeMeses.has(r.PeriodoLabel)
        );
    }
    function getFilteredTopArt()    {
        return TOP_ART.filter(r =>
            active.has(r.NombreSucursal) && activeMeses.has(r.PeriodoLabel)
        );
    }
    function getFilteredLineasCat() {
        return LINEAS_CAT.filter(r =>
            active.has(r.NombreSucursal) && activeMeses.has(r.PeriodoLabel)
        );
    }
    function getFilteredFabricantes() {
        return FABRICANTES.filter(r =>
            active.has(r.NombreSucursal) && activeMeses.has(r.PeriodoLabel)
        );
    }
    function salesColor(t){
        const clamped = Math.max(0, Math.min(1, t));
        const c1 = [224,224,224], c2 = [45,45,45];
        const r = Math.round(c1[0] + (c2[0]-c1[0])*clamped);
        const g = Math.round(c1[1] + (c2[1]-c1[1])*clamped);
        const b = Math.round(c1[2] + (c2[2]-c1[2])*clamped);
        return `rgba(${r},${g},${b},0.92)`;
    }
    function brandGradientColor(t, alpha){
        const clamped = Math.max(0, Math.min(1, t));
        const c1 = [224,224,224], c2 = [45,45,45];
        const r = Math.round(c1[0] + (c2[0]-c1[0])*clamped);
        const g = Math.round(c1[1] + (c2[1]-c1[1])*clamped);
        const b = Math.round(c1[2] + (c2[2]-c1[2])*clamped);
        return `rgba(${r},${g},${b},${alpha})`;
    }
    function brandGradientColors(values, alpha){
        const max = Math.max(...values, 0);
        const min = Math.min(...values, 0);
        return values.map(v => {
            const t = max>min ? (v-min)/(max-min) : 1;
            return brandGradientColor(0.15 + t*0.85, alpha);
        });
    }
    function getDates()      { return [...new Set(RAW.map(r => r.FechaStr))].sort(); }
    function getActiveSucs() { return [...active].sort(); }
    function aggByDate(data){
        const m = {};
        data.forEach(r => {
            if(!m[r.FechaStr]) m[r.FechaStr] = {unidades:0,ventas:0,utilidad:0,tickets:0};
            m[r.FechaStr].unidades  += r.unidades;
            m[r.FechaStr].ventas    += r.ventas;
            m[r.FechaStr].utilidad  += r.utilidad;
            m[r.FechaStr].tickets   += r.tickets;
        });
        return m;
    }
    function aggBySucursal(data){
        const m = {};
        data.forEach(r => {
            if(!m[r.NombreSucursal]) m[r.NombreSucursal] = {NombreSucursal:r.NombreSucursal, unidades:0, ventas:0, utilidad:0, tickets:0};
            m[r.NombreSucursal].unidades += r.unidades;
            m[r.NombreSucursal].ventas   += r.ventas;
            m[r.NombreSucursal].utilidad += r.utilidad;
            m[r.NombreSucursal].tickets  += r.tickets;
        });
        return Object.values(m);
    }
    function dc(id){ if(charts[id]){ charts[id].destroy(); delete charts[id]; } }
    window.changeLineMetric = function(metric){
        currentLineMetric = metric;
        document.querySelectorAll('[id^="btn-m-"]').forEach(b => b.classList.remove('active'));
        document.getElementById('btn-m-' + metric).classList.add('active');
        updateLineas();
    };
    window.changeLineMetricActual = function(metric){
        currentLineMetricActual = metric;
        document.querySelectorAll('[id^="btn-ma-"]').forEach(b => b.classList.remove('active'));
        document.getElementById('btn-ma-' + metric).classList.add('active');
        updateLineasActual();
    };
    window.changeHistMetric = function(metric){
        currentHistMetric = metric;
        document.querySelectorAll('[id^="btn-h-"]').forEach(b => b.classList.remove('active'));
        document.getElementById('btn-h-' + metric).classList.add('active');
        updateHistorico();
    };
    window.changeSucMetric = function(metric){
        currentSucMetric = metric;
        document.querySelectorAll('[id^="btn-hs-"]').forEach(b => b.classList.remove('active'));
        document.getElementById('btn-hs-' + metric).classList.add('active');
        updateChartSucursalActual();
    };
    function pronosticoData(){
        return PRONOSTICO.slice().sort((a,b) => a.ClaveSucursal - b.ClaveSucursal);
    }
    function updatePronostico(){
        updateTablaPronostico();
        updateChartPronostico();
        fitTables();
    }
    function updateTablaPronostico(){
        const data  = pronosticoData();
        const tbody = document.getElementById('tabla-pronostico');
        if(!data.length){
            tbody.innerHTML = '<tr><td colspan="6" class="empty">Ningún canal de venta tiene ventas registradas en el mes en curso.</td></tr>';
            ['p-unidades','p-ventas','p-util','p-pronostico'].forEach(id => { document.getElementById(id).textContent = '—'; });
            document.getElementById('p-util-sub').textContent = 'Margen: —';
            document.getElementById('p-pronostico-sub').textContent = 'Suma de pronósticos';
            return;
        }
        let tU=0, tV=0, tUt=0, tP=0;
        const rows = data.map(r => {
            tU += r.unidadesActual; tV += r.ventasActual; tUt += r.utilidadActual;
            const hayPron = r.pronostico !== null && r.pronostico !== undefined;
            if(hayPron) tP += r.pronostico;
            const mgCls = r.margen>=0.50 ? 'hi' : 'mi';
            return `<tr>
                <td data-label="Canal de Venta"><b>${r.NombreSucursal}</b></td>
                <td class="r" data-label="Unidades">${fN(r.unidadesActual)}</td>
                <td class="r" data-label="Venta $">${fF(r.ventasActual)}</td>
                <td class="r" data-label="Utilidad">${fF(r.utilidadActual)}</td>
                <td class="r" data-label="Margen"><span class="pill ${mgCls}">${fP(r.margen)}</span></td>
                <td class="r" data-label="Pronóstico de cierre"><b>${hayPron ? fF(r.pronostico) : '—'}</b></td>
            </tr>`;
        }).join('');
        const totMg = tV>0 ? tUt/tV : 0;
        tbody.innerHTML = rows + `<tr class="total-row">
            <td data-label=""><b>TOTAL GENERAL</b></td>
            <td class="r" data-label="Unidades">${fN(tU)}</td>
            <td class="r" data-label="Venta $">${fF(tV)}</td>
            <td class="r" data-label="Utilidad">${fF(tUt)}</td>
            <td class="r" data-label="Margen"><span class="pill ${totMg>=.50?'hi':'mi'}">${fP(totMg)}</span></td>
            <td class="r" data-label="Pronóstico de cierre"><b>${fF(tP)}</b></td>
        </tr>`;
        document.getElementById('p-unidades').textContent   = fN(tU);
        document.getElementById('p-ventas').textContent     = fF(tV);
        document.getElementById('p-util').textContent       = fF(tUt);
        document.getElementById('p-util-sub').textContent   = 'Margen: ' + fP(totMg);
        document.getElementById('p-pronostico').textContent = fF(tP);
        document.getElementById('p-pronostico-sub').textContent =
            tP > 0 ? `Avance real: ${fP(tV/tP)} de lo proyectado` : 'Suma de pronósticos';
    }
    function updateChartPronostico(){
        dc('pronostico');
        const data = pronosticoData();
        if(!data.length) return;
        document.getElementById('pron-chart-wrap').style.height =
            Math.max(360, Math.min(620, 340 + data.length * 6)) + 'px';
        const maxVal = Math.max(
            ...data.map(r => Math.max(r.ventasActual, r.pronostico || 0))
        ) || 0;
        charts['pronostico'] = new Chart(document.getElementById('chart-pronostico'), {
            data: {
                labels: data.map(r => r.NombreSucursal),
                datasets: [
                    {
                        type: 'bar',
                        label: 'Ventas acumuladas',
                        data: data.map(r => r.ventasActual),
                        backgroundColor: '#808080cc',
                        borderColor: '#5A5A5A',
                        borderWidth: 1,
                        borderRadius: 4,
                        order: 2,
                    },
                    {
                        type: 'line',
                        label: 'Pronóstico de cierre',
                        data: data.map(r => r.pronostico),
                        borderColor: '#aa7300',
                        backgroundColor: 'transparent',
                        borderWidth: 2,
                        pointRadius: 4,
                        pointHoverRadius: 6,
                        tension: 0.15,
                        fill: false,
                        order: 1,
                        datalabels: { display: false },
                    }
                ]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                interaction: { mode: 'index', intersect: false },
                plugins: {
                    legend: {
                        display: true, position: 'top', onClick: null,
                        labels: {
                            boxWidth: 9, boxHeight: 9, usePointStyle: true, pointStyle: 'circle',
                            padding: 12, color: '#616161',
                            font: { family: "'Segoe UI', sans-serif", size: 10, weight: '600' }
                        }
                    },
                    datalabels: { display: false },
                    tooltip: {
                        ...PREMIUM_TOOLTIP_OPTS,
                        filter: ctx => ctx.raw !== null && ctx.raw !== undefined,
                        callbacks: {
                            label: ctx => ` ${ctx.dataset.label}: ${fF(ctx.raw)}`,
                            footer: items => {
                                if(!items.length) return '';
                                const r = data[items[0].dataIndex];
                                if(!(r.pronostico > 0)) return '';
                                const avance = r.ventasActual / r.pronostico;
                                return `Avance real: ${fP(avance)} · ${r.diasTranscurridos} de ${r.diasOperativosMes} días`;
                            }
                        },
                        footerColor: '#3D3D3D',
                        footerFont: { family: "'Segoe UI', sans-serif", size: 11, weight: 'bold' }
                    }
                },
                scales: {
                    x: {
                        grid: { display: false },
                        ticks: { font: { size: 9 }, color: '#858585', maxRotation: 60, minRotation: 45, autoSkip: false }
                    },
                    y: {
                        grid: { color: '#f4f4f4' },
                        suggestedMax: maxVal * 1.1,
                        ticks: { font: { size: 9 }, color: '#858585', callback: v => fM(v) }
                    }
                }
            }
        });
    }
    function update(){
        const data       = getFiltered();
        const dates      = getDates();
        const byDate     = aggByDate(data);
        const activeSucs = getActiveSucs();
        const multi      = activeSucs.length > 1;
        const totV  = data.reduce((a,r) => a+r.ventas,   0);
        const totU  = data.reduce((a,r) => a+r.utilidad, 0);
        const totUn = data.reduce((a,r) => a+r.unidades, 0);
        const totTk = data.reduce((a,r) => a + r.tickets, 0);
        document.getElementById('k-uni').textContent      = fN(totUn);
        document.getElementById('k-ventas').textContent   = fF(totV);
        document.getElementById('k-ventas-s').textContent = activeSucs.length + ' canal(es)';
        document.getElementById('k-util').textContent     = fF(totU);
        document.getElementById('k-util-sub').textContent = 'Margen: ' + fP(totV>0 ? totU/totV : 0);
        document.getElementById('k-tkt').textContent       = fN(totTk);
        document.getElementById('k-tkt-sub').textContent   = 'Volumen de ventas';
        document.getElementById('k-vtkt').textContent      = totTk>0 ? fF(totV/totTk) : '—';
        document.getElementById('k-vtkt-sub').textContent  = 'Ventas $ ÷ Tickets';
        document.getElementById('k-utkt').textContent      = totTk>0 ? (totUn/totTk).toFixed(1) : '—';
        document.getElementById('k-utkt-sub').textContent  = 'Unidades ÷ Tickets';
        updateTabla(dates, byDate);
        updateChart('ventas',  dates, byDate, activeSucs, multi);
        updateChart('tickets', dates, byDate, activeSucs, multi);
        fitTables();
    }
    function updateTabla(dates, byDate){
        const tbody = document.getElementById('tabla-body');
        const valid = dates.filter(d => byDate[d]);
        if(!valid.length){
            tbody.innerHTML = '<tr><td colspan="7" class="empty">Selecciona canales de venta para mapear la grilla.</td></tr>';
            return;
        }
        let tUn=0,tV=0,tU=0,tTk=0;
        const rows = valid.map(d => {
            const r  = byDate[d];
            const mg = r.ventas>0 ? r.utilidad/r.ventas : 0;
            const pill = mg>=.50 ? 'hi' : 'mi';
            const dn = DAYS[new Date(d+'T12:00:00').getDay()];
            const mAbr = MESES_ABR[parseInt(d.slice(5,7),10)-1];
            tUn+=r.unidades; tV+=r.ventas; tU+=r.utilidad; tTk+=r.tickets;
            return `<tr><td class="date" data-label="Fecha">${d.slice(8)} ${mAbr}</td><td class="dayname" data-label="Día">${dn}</td><td class="r" data-label="Unidades">${fN(r.unidades)}</td><td class="r" data-label="Ventas $"><b>${fF(r.ventas)}</b></td><td class="r" data-label="Utilidad">${fF(r.utilidad)}</td><td class="r" data-label="Margen"><span class="pill ${pill}">${fP(mg)}</span></td><td class="r" data-label="Tickets"><b>${fN(r.tickets)}</b></td></tr>`;
        }).join('');
        const totMg = tV>0 ? tU/tV : 0;
        tbody.innerHTML = rows + `<tr class="total-row"><td data-label="" colspan="2"><b>TOTAL PERÍODO</b></td><td class="r" data-label="Unidades">${fN(tUn)}</td><td class="r" data-label="Ventas $">${fF(tV)}</td><td class="r" data-label="Utilidad">${fF(tU)}</td><td class="r" data-label="Margen"><span class="pill ${totMg>=.50?'hi':'mi'}">${fP(totMg)}</span></td><td class="r" data-label="Tickets">${fN(tTk)}</td></tr>`;
    }
    function updateChartGeneric(chartId, field, sourceData, dates, byDate, activeSucs, multi){
        dc(chartId);
        const canvasEl = document.getElementById('chart-'+chartId);
        if(!canvasEl) return;
        const labels    = dates.map(d => d.slice(8));
        const isVentas  = field === 'ventas';
        const monoColor = isVentas ? '#5A5A5A' : '#2196A8';
        const datasets  = multi
            ? activeSucs.map(s => ({
                label: s,
                data: dates.map(d => { const r = sourceData.find(x => x.FechaStr===d && x.NombreSucursal===s); return r ? r[field] : 0; }),
                borderColor: SC[s]||'#888', borderWidth:1.6,
                pointRadius:3.5, pointHitRadius:20, pointHoverRadius:6, tension:0.12, fill:false
              }))
            : [{
                label: isVentas ? 'Ventas c/Desc' : 'Cantidad de Tickets',
                data: dates.map(d => byDate[d] ? byDate[d][field] : 0),
                borderColor: monoColor, backgroundColor: monoColor+'08', borderWidth:2.2,
                pointRadius:3.5, pointHitRadius:20, pointHoverRadius:6, tension:0.12, fill:true
              }];
        charts[chartId] = new Chart(canvasEl, {
            type:'line', data:{labels, datasets},
            options:{
                responsive:true, maintainAspectRatio:false,
                interaction:{mode:'index', intersect:false},
                plugins:{
                    legend:{display:multi, position:'top', onClick:null,
                        labels:{boxWidth:8,boxHeight:8,usePointStyle:true,pointStyle:'circle',padding:12,color:'#616161',font:{family:"'Segoe UI', sans-serif",size:10,weight:'600'}}},
                    datalabels:{display:false},
                    tooltip:{...PREMIUM_TOOLTIP_OPTS, callbacks:{
                        label: ctx => ' ' + (ctx.dataset.label ? ctx.dataset.label + ': ' : '')
                                          + (isVentas ? fF(ctx.raw) : fN(ctx.raw))
                    }}
                },
                scales:{
                    x:{grid:{display:false}, ticks:{font:{size:10},color:'#858585'}},
                    y:{grid:{color:'#f4f4f4'}, ticks:{font:{size:10},color:'#858585', callback: isVentas ? v=>fM(v) : v=>v}}
                }
            }
        });
    }
    function updateChart(field, dates, byDate, activeSucs, multi){
        updateChartGeneric(field, field, RAW, dates, byDate, activeSucs, multi);
    }
    function updateLineas(){
        dc('lineas');
        const tbody    = document.getElementById('tabla-lineas');
        const filtered = getFilteredLineas();
        const lineMap  = {};
        let totalUnidades=0, totalVentas=0, totalUtilidad=0;
        filtered.forEach(r => {
            if(!lineMap[r.Línea]) lineMap[r.Línea] = {Línea:r.Línea, unidades:0, ventas:0, utilidad:0};
            lineMap[r.Línea].unidades += r.unidades;
            lineMap[r.Línea].ventas   += r.ventas;
            lineMap[r.Línea].utilidad += r.utilidad;
            totalUnidades += r.unidades; totalVentas += r.ventas; totalUtilidad += r.utilidad;
        });
        const sorted = Object.values(lineMap).sort((a,b) => b[currentLineMetric]-a[currentLineMetric]);
        if(!sorted.length){
            tbody.innerHTML = '<tr><td colspan="5" class="empty">Selecciona canales de venta para desplegar líneas.</td></tr>';
            document.getElementById('lineas-chart-wrap').style.height = '220px';
            fitTables();
            return;
        }
        const rows = sorted.map(l => {
            const mg = l.ventas>0 ? l.utilidad/l.ventas : 0;
            return `<tr><td data-label="Línea"><span style="display:inline-block;width:8px;height:8px;border-radius:4px;background:${LC[l.Línea]||'#888'};margin-right:7px"></span>${l.Línea}</td><td class="r" data-label="Unidades">${fN(l.unidades)}</td><td class="r" data-label="Ventas $"><b>${fF(l.ventas)}</b></td><td class="r" data-label="Utilidad">${fF(l.utilidad)}</td><td class="r" data-label="Margen"><span class="pill ${mg>=.50?'hi':'mi'}">${fP(mg)}</span></td></tr>`;
        }).join('');
        const totalMargen = totalVentas>0 ? totalUtilidad/totalVentas : 0;
        tbody.innerHTML = rows + `<tr class="total-row"><td data-label=""><b>TOTAL ACUMULADO</b></td><td class="r" data-label="Unidades">${fN(totalUnidades)}</td><td class="r" data-label="Ventas $">${fF(totalVentas)}</td><td class="r" data-label="Utilidad">${fF(totalUtilidad)}</td><td class="r" data-label="Margen"><span class="pill ${totalMargen>=.50?'hi':'mi'}">${fP(totalMargen)}</span></td></tr>`;
        const maxVal = Math.max(...sorted.map(l => l[currentLineMetric]))||0;
        document.getElementById('lineas-chart-wrap').style.height =
            Math.max(220, sorted.length*34 + 70) + 'px';
        charts['lineas'] = new Chart(document.getElementById('chart-lineas'), {
            type:'bar',
            data:{labels:sorted.map(l=>l.Línea), datasets:[{data:sorted.map(l=>l[currentLineMetric]), backgroundColor:sorted.map(l=>LC[l.Línea]||'#888'), borderRadius:5, barThickness:14}]},
            options:{
                indexAxis:'y', responsive:true, maintainAspectRatio:false,
                layout:{padding:{right:28}},
                plugins:{
                    legend:{display:false},
                    tooltip:{...PREMIUM_TOOLTIP_OPTS, callbacks:{label:ctx=>(currentLineMetric==='ventas'||currentLineMetric==='utilidad')?' '+fF(ctx.raw):' '+fN(ctx.raw)+' uds'}},
                    datalabels:{display:true,anchor:'end',align:'end',color:'#4d4d4d',font:{weight:'600',size:9.5},
                        formatter:(value,ctx)=>{const t=ctx.chart.data.datasets[ctx.datasetIndex].data.reduce((a,b)=>a+b,0); return fP(t>0?value/t:0);}}
                },
                scales:{
                    x:{grid:{display:false},border:{display:false},suggestedMax:maxVal*1.15,ticks:{font:{size:9},color:'#858585',callback:currentLineMetric==='unidades'?v=>fN(v):v=>fM(v)}},
                    y:{grid:{display:false},border:{display:false},ticks:{font:{size:10,weight:'600'},color:'#212121'}}
                }
            }
        });
        fitTables();
    }
    function updateLineasActual(){
        dc('lineasactual');
        const tbody    = document.getElementById('tabla-lineas-actual');
        const filtered = getFilteredLineasActual();
        const lineMap  = {};
        let totalUnidades=0, totalVentas=0, totalUtilidad=0;
        filtered.forEach(r => {
            if(!lineMap[r.Línea]) lineMap[r.Línea] = {Línea:r.Línea, unidades:0, ventas:0, utilidad:0};
            lineMap[r.Línea].unidades += r.unidades;
            lineMap[r.Línea].ventas   += r.ventas;
            lineMap[r.Línea].utilidad += r.utilidad;
            totalUnidades += r.unidades; totalVentas += r.ventas; totalUtilidad += r.utilidad;
        });
        const sorted = Object.values(lineMap).sort((a,b) => b[currentLineMetricActual]-a[currentLineMetricActual]);
        if(!sorted.length){
            tbody.innerHTML = '<tr><td colspan="5" class="empty">Selecciona canales de venta para desplegar líneas.</td></tr>';
            document.getElementById('lineasactual-chart-wrap').style.height = '220px';
            return;
        }
        const rows = sorted.map(l => {
            const mg = l.ventas>0 ? l.utilidad/l.ventas : 0;
            return `<tr><td data-label="Línea"><span style="display:inline-block;width:8px;height:8px;border-radius:4px;background:${LC[l.Línea]||'#888'};margin-right:7px"></span>${l.Línea}</td><td class="r" data-label="Unidades">${fN(l.unidades)}</td><td class="r" data-label="Ventas $"><b>${fF(l.ventas)}</b></td><td class="r" data-label="Utilidad">${fF(l.utilidad)}</td><td class="r" data-label="Margen"><span class="pill ${mg>=.50?'hi':'mi'}">${fP(mg)}</span></td></tr>`;
        }).join('');
        const totalMargen = totalVentas>0 ? totalUtilidad/totalVentas : 0;
        tbody.innerHTML = rows + `<tr class="total-row"><td data-label=""><b>TOTAL MES EN CURSO</b></td><td class="r" data-label="Unidades">${fN(totalUnidades)}</td><td class="r" data-label="Ventas $">${fF(totalVentas)}</td><td class="r" data-label="Utilidad">${fF(totalUtilidad)}</td><td class="r" data-label="Margen"><span class="pill ${totalMargen>=.50?'hi':'mi'}">${fP(totalMargen)}</span></td></tr>`;
        const maxVal = Math.max(...sorted.map(l => l[currentLineMetricActual]))||0;
        document.getElementById('lineasactual-chart-wrap').style.height =
            Math.max(220, sorted.length*34 + 70) + 'px';
        charts['lineasactual'] = new Chart(document.getElementById('chart-lineasactual'), {
            type:'bar',
            data:{labels:sorted.map(l=>l.Línea), datasets:[{data:sorted.map(l=>l[currentLineMetricActual]), backgroundColor:sorted.map(l=>LC[l.Línea]||'#888'), borderRadius:5, barThickness:14}]},
            options:{
                indexAxis:'y', responsive:true, maintainAspectRatio:false,
                layout:{padding:{right:28}},
                plugins:{
                    legend:{display:false},
                    tooltip:{...PREMIUM_TOOLTIP_OPTS, callbacks:{label:ctx=>(currentLineMetricActual==='ventas'||currentLineMetricActual==='utilidad')?' '+fF(ctx.raw):' '+fN(ctx.raw)+' uds'}},
                    datalabels:{display:true,anchor:'end',align:'end',color:'#4d4d4d',font:{weight:'600',size:9.5},
                        formatter:(value,ctx)=>{const t=ctx.chart.data.datasets[ctx.datasetIndex].data.reduce((a,b)=>a+b,0); return fP(t>0?value/t:0);}}
                },
                scales:{
                    x:{grid:{display:false},border:{display:false},suggestedMax:maxVal*1.15,ticks:{font:{size:9},color:'#858585',callback:currentLineMetricActual==='unidades'?v=>fN(v):v=>fM(v)}},
                    y:{grid:{display:false},border:{display:false},ticks:{font:{size:10,weight:'600'},color:'#212121'}}
                }
            }
        });
    }
    function updateTablaSucursalActual(){
        const tbody = document.getElementById('tabla-sucursal-actual');
        const data  = aggBySucursal(getFiltered()).sort((a,b) => b.ventas - a.ventas);
        if(!data.length){
            tbody.innerHTML = '<tr><td colspan="6" class="empty">Selecciona canales de venta para ver el comparativo.</td></tr>';
            return;
        }
        let tUn=0,tV=0,tU=0,tTk=0;
        const rows = data.map(r => {
            const mg = r.ventas>0 ? r.utilidad/r.ventas : 0;
            tUn+=r.unidades; tV+=r.ventas; tU+=r.utilidad; tTk+=r.tickets;
            return `<tr><td data-label="Canal de Venta"><span style="display:inline-block;width:8px;height:8px;border-radius:4px;background:${SC[r.NombreSucursal]||'#888'};margin-right:7px"></span><b>${r.NombreSucursal}</b></td><td class="r" data-label="Unidades">${fN(r.unidades)}</td><td class="r" data-label="Ventas $"><b>${fF(r.ventas)}</b></td><td class="r" data-label="Utilidad">${fF(r.utilidad)}</td><td class="r" data-label="Margen"><span class="pill ${mg>=.50?'hi':'mi'}">${fP(mg)}</span></td><td class="r" data-label="Tickets">${fN(r.tickets)}</td></tr>`;
        }).join('');
        const totMg = tV>0 ? tU/tV : 0;
        tbody.innerHTML = rows + `<tr class="total-row"><td data-label=""><b>TOTAL MES EN CURSO</b></td><td class="r" data-label="Unidades">${fN(tUn)}</td><td class="r" data-label="Ventas $">${fF(tV)}</td><td class="r" data-label="Utilidad">${fF(tU)}</td><td class="r" data-label="Margen"><span class="pill ${totMg>=.50?'hi':'mi'}">${fP(totMg)}</span></td><td class="r" data-label="Tickets">${fN(tTk)}</td></tr>`;
    }
    function updateChartSucursalActual(){
        dc('sucursalactual');
        const data = aggBySucursal(getFiltered()).sort((a,b) => b[currentSucMetric]-a[currentSucMetric]);
        const wrap = document.getElementById('sucursalactual-chart-wrap');
        if(!data.length){ if(wrap) wrap.style.height = '220px'; return; }
        const isMoneda  = currentSucMetric==='ventas' || currentSucMetric==='utilidad';
        const margenes  = data.map(r => r.ventas>0 ? +(r.utilidad/r.ventas*100).toFixed(2) : 0);
        const metricVals = data.map(r => r[currentSucMetric]);
        const barFill    = brandGradientColors(metricVals, 0.85);
        const barBorder  = brandGradientColors(metricVals, 1);
        if(wrap) wrap.style.height = Math.max(260, data.length*40 + 90) + 'px';
        charts['sucursalactual'] = new Chart(document.getElementById('chart-sucursalactual'), {
            data: {
                labels: data.map(r => r.NombreSucursal),
                datasets: [
                    {
                        type:'bar',
                        label: currentSucMetric==='ventas'?'Ventas c/Desc':currentSucMetric==='utilidad'?'Utilidad':currentSucMetric==='unidades'?'Unidades':'Tickets',
                        data: data.map(r => r[currentSucMetric]),
                        backgroundColor: barFill,
                        borderColor: barBorder,
                        borderWidth:1, borderRadius:5, yAxisID:'y'
                    },
                    {
                        type:'line', label:'Margen %', yAxisID:'y2',
                        data: margenes, borderColor:'#aa7300', backgroundColor:'transparent',
                        borderWidth:2, pointRadius:4, pointHoverRadius:6, tension:0.15,
                        datalabels:{display:false}
                    }
                ]
            },
            options:{
                responsive:true, maintainAspectRatio:false,
                interaction:{mode:'index', intersect:false},
                plugins:{
                    legend:{display:true, position:'top', onClick:null,
                        labels:{boxWidth:8,boxHeight:8,usePointStyle:true,pointStyle:'circle',padding:10,color:'#616161',font:{family:"'Segoe UI', sans-serif",size:10,weight:'600'}}},
                    datalabels:{display:false},
                    tooltip:{...PREMIUM_TOOLTIP_OPTS, callbacks:{
                        label: ctx => {
                            if(ctx.dataset.yAxisID==='y2') return ` Margen: ${ctx.raw.toFixed(2)}%`;
                            return ` ${ctx.dataset.label}: ${isMoneda?fF(ctx.raw):fN(ctx.raw)}`;
                        }
                    }}
                },
                scales:{
                    x:{grid:{display:false}, ticks:{font:{size:9},color:'#858585',maxRotation:45}},
                    y:{position:'left', grid:{color:'#f4f4f4'}, ticks:{font:{size:9},color:'#858585', callback: v => isMoneda?fM(v):fN(v)}},
                    y2:{position:'right', grid:{display:false}, ticks:{font:{size:9},color:'#aa7300', callback: v => v.toFixed(1)+'%'}}
                }
            }
        });
    }
    function updateResumenActualExtra(){
        updateTablaSucursalActual();
        updateChartSucursalActual();
        updateLineasActual();
        fitTables();
    }
    function updateKPIsAcumulado(){
        const data  = getFilteredHist();
        const totU  = data.reduce((a,r) => a+r.unidades, 0);
        const totV  = data.reduce((a,r) => a+r.ventas,   0);
        const totUt = data.reduce((a,r) => a+r.utilidad, 0);
        const periodosCons    = periodosConsolidadosEn(data);
        const huboConsolidado = periodosCons.length > 0;
        const totTkCons = periodosCons.reduce((a,p) => a + TICKETS_CONS[p], 0);
        const totTk     = data.reduce((a,r) => a + r.tickets, 0) + totTkCons;
        const marcaCons = huboConsolidado ? ' *' : '';
        const periodosEnUso = new Set(data.map(r => r.PeriodoLabel));
        const baseProm = huboConsolidado
            ? HISTORICO.filter(r => periodosEnUso.has(r.PeriodoLabel))
            : data;
        const promV = baseProm.reduce((a,r) => a+r.ventas,   0);
        const promU = baseProm.reduce((a,r) => a+r.unidades, 0);
        const promT = baseProm.reduce((a,r) => a+r.tickets,  0) + totTkCons;
        const sufijoBase = huboConsolidado ? ' · total del negocio' : '';
        const notaEl = document.getElementById('nota-consolidado');
        if(notaEl){
            notaEl.style.display = HAY_CONSOLIDADO ? 'block' : 'none';
            notaEl.textContent   = HAY_CONSOLIDADO ? NOTA_CONSOLIDADO : '';
        }
        document.getElementById('ka-uni').textContent      = fN(totU);
        document.getElementById('ka-ventas').textContent   = fF(totV);
        document.getElementById('ka-ventas-s').textContent = getActiveSucs().length + ' canal(es) · ' + activeMeses.size + ' período(s)';
        document.getElementById('ka-util').textContent     = fF(totUt);
        document.getElementById('ka-util-sub').textContent = 'Margen: ' + fP(totV>0 ? totUt/totV : 0);
        document.getElementById('ka-tkt').textContent       = fN(totTk) + marcaCons;
        document.getElementById('ka-tkt-sub').textContent   = 'Volumen de ventas';
        document.getElementById('ka-vtkt').textContent      = promT>0 ? fF(promV/promT) + marcaCons : '—';
        document.getElementById('ka-vtkt-sub').textContent  = 'Ventas $ ÷ Tickets' + sufijoBase;
        document.getElementById('ka-utkt').textContent      = promT>0 ? (promU/promT).toFixed(1) + marcaCons : '—';
        document.getElementById('ka-utkt-sub').textContent  = 'Unidades ÷ Tickets' + sufijoBase;
    }
    function updateAcumulado(){
        updateKPIsAcumulado();
        updateHistorico();
        updateHistComparativos();
        updateLineas();
        fitTables();
    }
    function updateHistorico(){
        dc('historico');
        const tbody    = document.getElementById('tabla-historico');
        const filtered = getFilteredHist();
        if(!filtered.length){
            tbody.innerHTML = '<tr><td colspan="6" class="empty">Selecciona canales de venta y meses para ver el histórico.</td></tr>';
            const notaVacia = document.getElementById('nota-tendencia');
            if(notaVacia) notaVacia.style.display = 'none';
            fitTables();
            return;
        }
        const periodoMap   = {};
        const periodoOrder = [];
        filtered.forEach(r => {
            if(!periodoMap[r.PeriodoLabel]){
                periodoMap[r.PeriodoLabel] = {PeriodoLabel:r.PeriodoLabel, Año:r.Año, MesNum:r.MesNum,
                    unidades:0, ventas:0, utilidad:0, tickets:0};
                periodoOrder.push(r.PeriodoLabel);
            }
            periodoMap[r.PeriodoLabel].unidades  += r.unidades;
            periodoMap[r.PeriodoLabel].ventas    += r.ventas;
            periodoMap[r.PeriodoLabel].utilidad  += r.utilidad;
            periodoMap[r.PeriodoLabel].tickets   += r.tickets;
        });
        const sorted = periodoOrder
            .map(p => periodoMap[p])
            .sort((a,b) => a.Año!==b.Año ? a.Año-b.Año : a.MesNum-b.MesNum);
        let huboConsolidado = false;
        sorted.forEach(r => {
            if(TICKETS_CONS[r.PeriodoLabel] !== undefined){
                r.tickets       = TICKETS_CONS[r.PeriodoLabel];
                r.consolidado   = true;
                huboConsolidado = true;
            }
        });
        let tUn=0,tV=0,tU=0,tTk=0;
        sorted.forEach(r => { tUn+=r.unidades; tV+=r.ventas; tU+=r.utilidad; tTk+=r.tickets; });
        const totMg = tV>0 ? tU/tV : 0;
        const rows = sorted.map(r => {
            const mg = r.ventas>0 ? r.utilidad/r.ventas : 0;
            const tkCell = `<b>${fN(r.tickets)}</b>`
                + (r.consolidado ? ' <span title="Sin detalle por canal de venta: el total corresponde a todos los canales" style="color:#c4c4c4;font-weight:700">*</span>' : '');
            return `<tr><td class="date" data-label="Período">${r.PeriodoLabel}</td><td class="r" data-label="Unidades">${fN(r.unidades)}</td><td class="r" data-label="Ventas $"><b>${fF(r.ventas)}</b></td><td class="r" data-label="Utilidad">${fF(r.utilidad)}</td><td class="r" data-label="Margen"><span class="pill ${mg>=.50?'hi':'mi'}">${fP(mg)}</span></td><td class="r" data-label="Tickets">${tkCell}</td></tr>`;
        }).join('');
        const totTkTexto = fN(tTk) + (huboConsolidado ? ' *' : '');
        const notaPieCons = '* No existe desglose de tickets por canal de venta. Se muestra acumulado.';
        const pieEstilo = 'font-size:.65rem;color:#a3a3a3;text-align:right;border:none;padding-top:4px';
        tbody.innerHTML = rows + `<tr class="total-row"><td data-label=""><b>TOTAL HISTÓRICO</b></td><td class="r" data-label="Unidades">${fN(tUn)}</td><td class="r" data-label="Ventas $">${fF(tV)}</td><td class="r" data-label="Utilidad">${fF(tU)}</td><td class="r" data-label="Margen"><span class="pill ${totMg>=.50?'hi':'mi'}">${fP(totMg)}</span></td><td class="r" data-label="Tickets">${totTkTexto}</td></tr>`
            + (huboConsolidado ? `<tr><td colspan="6" style="${pieEstilo}">${notaPieCons}</td></tr>` : '');
        const activeSucs = getActiveSucs();
        const multi      = activeSucs.length > 1;
        const isMoneda   = currentHistMetric==='ventas' || currentHistMetric==='utilidad';
        const labels     = sorted.map(r => r.PeriodoLabel);
        const esCons = r => currentHistMetric==='tickets' && r.consolidado === true;
        const hayCons = sorted.some(esCons);
        const notaTend = document.getElementById('nota-tendencia');
        if(notaTend){
            notaTend.style.display = hayCons ? 'block' : 'none';
            notaTend.textContent   = hayCons ? NOTA_TENDENCIA : '';
        }
        const canalesHistorico = [...new Set(HISTORICO.map(x => x.NombreSucursal))];
        const sinFiltroCanal = canalesHistorico.every(x => active.has(x));
        const crecimientos = sorted.map(r => {
            let prevMes = r.MesNum - 1, prevAnio = r.Año;
            if(prevMes === 0){ prevMes = 12; prevAnio -= 1; }
            const prevLabel = labelPeriodo(prevAnio, prevMes);
            const prevEsCons = currentHistMetric==='tickets' && TICKETS_CONS[prevLabel] !== undefined;
            if(prevEsCons !== esCons(r) && !sinFiltroCanal) return null;
            let prevTotal;
            if(prevEsCons){
                prevTotal = TICKETS_CONS[prevLabel];
            } else {
                prevTotal = HISTORICO
                    .filter(x => active.has(x.NombreSucursal) && x.Año===prevAnio && x.MesNum===prevMes)
                    .reduce((a,x) => a + x[currentHistMetric], 0);
            }
            if(prevTotal <= 0) return null;
            return +(((r[currentHistMetric] - prevTotal) / prevTotal) * 100).toFixed(2);
        });
        const barDatasets = multi
            ? activeSucs.map(s => {
                const color = SC[s]||'#888';
                return {
                    type:'bar', label:s, yAxisID:'y',
                    data: sorted.map(p => {
                        if(esCons(p)) return null;
                        return filtered
                            .filter(r => r.NombreSucursal===s && r.PeriodoLabel===p.PeriodoLabel)
                            .reduce((a,r) => a+r[currentHistMetric], 0);
                    }),
                    backgroundColor: color+'cc', borderColor: color,
                    borderWidth: 1, borderRadius: 3, stack: 'stack0',
                };
              })
            : [{
                type:'bar',
                label: currentHistMetric==='ventas'?'Ventas c/Desc':currentHistMetric==='utilidad'?'Utilidad':currentHistMetric==='unidades'?'Unidades':'Tickets',
                data: sorted.map(r => esCons(r) ? null : r[currentHistMetric]),
                backgroundColor:'#808080cc', borderColor:'#5A5A5A',
                borderWidth:1, borderRadius:4, yAxisID:'y',
              }];
        if(hayCons){
            barDatasets.push({
                type:'bar', label:'Tickets acumulados *', yAxisID:'y',
                data: sorted.map(r => esCons(r) ? r.tickets : null),
                backgroundColor:'#b5b5b5cc', borderColor:'#7A7A7A',
                borderWidth:1, borderRadius: multi ? 3 : 4,
                stack: multi ? 'stack0' : undefined,
            });
        }
        const crecimientoLine = {
            type:'line', label:'Crecimiento % (vs. mes anterior)', yAxisID:'y2',
            data: crecimientos,
            borderColor:'#aa7300', backgroundColor:'transparent',
            borderWidth:2, pointRadius:3.5, pointHoverRadius:6, tension:0.2,
            spanGaps:false,
            datalabels:{display:false}
        };
        charts['historico'] = new Chart(document.getElementById('chart-historico'), {
            data:{ labels, datasets:[...barDatasets, crecimientoLine] },
            options:{
                responsive:true, maintainAspectRatio:false,
                interaction:{mode:'index', intersect:false},
                plugins:{
                    legend:{display:true, position:'top', onClick:null,
                        labels:{boxWidth:8,boxHeight:8,usePointStyle:true,pointStyle:'circle',padding:10,color:'#616161',
                            font:{family:"'Segoe UI', sans-serif",size:10,weight:'600'}}},
                    datalabels:{display:false},
                    tooltip:{...PREMIUM_TOOLTIP_OPTS,
                        filter: ctx => !(ctx.dataset.yAxisID!=='y2' && (ctx.raw === null || ctx.raw === undefined)),
                        callbacks:{
                        label: ctx => {
                            if(ctx.dataset.yAxisID==='y2'){
                                if(ctx.raw === null || ctx.raw === undefined) return ' Crecimiento: sin dato del mes anterior';
                                const signo = ctx.raw >= 0 ? '+' : '';
                                return ` Crecimiento vs. mes anterior: ${signo}${ctx.raw.toFixed(2)}%`;
                            }
                            return isMoneda ? ` ${ctx.dataset.label}: ${fF(ctx.raw)}` : ` ${ctx.dataset.label}: ${fN(ctx.raw)}`;
                        }
                    }}
                },
                scales:{
                    x:{grid:{display:false}, ticks:{font:{size:9},color:'#858585',maxRotation:45}},
                    y:{stacked:true, position:'left', grid:{color:'#f4f4f4'}, ticks:{font:{size:9},color:'#858585', callback: v => isMoneda?fM(v):fN(v)}},
                    y2:{position:'right', grid:{display:false}, ticks:{font:{size:9},color:'#aa7300', callback: v => v.toFixed(1)+'%'}}
                }
            }
        });
        fitTables();
    }
    function totalVentasPeriodo(anio, mes){
        return HISTORICO
            .filter(r => active.has(r.NombreSucursal) && r.Año===anio && r.MesNum===mes)
            .reduce((a,r) => a + r.ventas, 0);
    }
    function labelPeriodo(anio, mes){
        return MESES_ABR[mes-1].charAt(0).toUpperCase() + MESES_ABR[mes-1].slice(1) + '-' + String(anio).slice(-2);
    }
    function renderComparativoVentas(canvasId, chartKey, labelRef, valRef, labelAct, valAct){
        dc(chartKey);
        const max = Math.max(valRef, valAct, 0);
        charts[chartKey] = new Chart(document.getElementById(canvasId), {
            type: 'bar',
            data: {
                labels: [labelRef, labelAct],
                datasets: [{
                    data: [valRef, valAct],
                    backgroundColor: ['#C4C4C4', '#5A5A5A'],
                    borderColor: ['#A8A8A8', '#3D3D3D'],
                    borderWidth: 1, borderRadius: 6, maxBarThickness: 90,
                }]
            },
            options: {
                responsive: true, maintainAspectRatio: false,
                plugins: {
                    legend: { display: false },
                    datalabels: {
                        display: true, anchor: 'end', align: 'end', offset: 2,
                        color: '#4d4d4d', font: { weight: '700', size: 11 },
                        formatter: v => fF(v)
                    },
                    tooltip: { ...PREMIUM_TOOLTIP_OPTS, callbacks: {
                        label: ctx => ` Ventas: ${fF(ctx.raw)}`
                    }}
                },
                scales: {
                    x: { grid: { display: false }, ticks: { font: { size: 10, weight: '600' }, color: '#616161' } },
                    y: { grid: { color: '#f4f4f4' }, suggestedMax: max > 0 ? max * 1.18 : 1, ticks: { font: { size: 9 }, color: '#858585', callback: v => fM(v) } }
                }
            }
        });
    }
    function baseHistoricoActual(){
        if(activeMeses.size === 1){
            const p   = [...activeMeses][0];
            const row = HISTORICO.find(r => r.PeriodoLabel === p);
            if(row) return { anio: row.Año, mes: row.MesNum, label: p };
        }
        const actualRow = HISTORICO.find(r => r.PeriodoLabel === CURRENT_PERIOD);
        if(actualRow) return { anio: actualRow.Año, mes: actualRow.MesNum, label: CURRENT_PERIOD };
        return null;
    }
    function updateHistComparativos(){
        const anioWrap = document.getElementById('histcomp-anio-wrap');
        const mesWrap  = document.getElementById('histcomp-mes-wrap');
        const anioVar  = document.getElementById('histcomp-anio-var');
        const mesVar   = document.getElementById('histcomp-mes-var');
        dc('histcompanio'); dc('histcompmes');
        const base = baseHistoricoActual();
        if(!base){
            anioWrap.innerHTML = '<p class="empty" style="padding:2rem 0">Sin datos del mes en curso.</p>';
            mesWrap.innerHTML  = '<p class="empty" style="padding:2rem 0">Sin datos del mes en curso.</p>';
            document.getElementById('histcomp-anio-title').textContent = 'Ventas: Mes en Curso vs. Mismo Mes Año Anterior';
            document.getElementById('histcomp-mes-title').textContent  = 'Ventas: Mes en Curso vs. Mes Anterior';
            anioVar.textContent = '—'; anioVar.className = 'pill';
            mesVar.textContent  = '—'; mesVar.className  = 'pill';
            return;
        }
        anioWrap.innerHTML = '<canvas id="chart-hist-comp-anio"></canvas>';
        mesWrap.innerHTML  = '<canvas id="chart-hist-comp-mes"></canvas>';
        const anioAct = base.anio, mesAct = base.mes, labelAct = base.label;
        const valActual = totalVentasPeriodo(anioAct, mesAct);
        const anioAnt      = anioAct - 1;
        const valAnioAnt   = totalVentasPeriodo(anioAnt, mesAct);
        const labelAnioAnt = labelPeriodo(anioAnt, mesAct);
        let mesPrev = mesAct - 1, anioPrev = anioAct;
        if(mesPrev === 0){ mesPrev = 12; anioPrev -= 1; }
        const valMesAnt   = totalVentasPeriodo(anioPrev, mesPrev);
        const labelMesAnt = labelPeriodo(anioPrev, mesPrev);
        document.getElementById('histcomp-anio-title').textContent = `Ventas: ${labelAct} vs. ${labelAnioAnt}`;
        document.getElementById('histcomp-mes-title').textContent  = `Ventas: ${labelAct} vs. ${labelMesAnt}`;
        renderComparativoVentas('chart-hist-comp-anio', 'histcompanio', labelAnioAnt, valAnioAnt, labelAct, valActual);
        renderComparativoVentas('chart-hist-comp-mes',  'histcompmes',  labelMesAnt,  valMesAnt,  labelAct, valActual);
        const varAnio = valAnioAnt > 0 ? ((valActual - valAnioAnt) / valAnioAnt * 100) : null;
        const varMes  = valMesAnt  > 0 ? ((valActual - valMesAnt)  / valMesAnt  * 100) : null;
        anioVar.textContent = varAnio === null ? 'Sin dato de referencia' : `${varAnio >= 0 ? '+' : ''}${varAnio.toFixed(2)}%`;
        anioVar.className   = 'pill ' + (varAnio === null ? '' : varAnio >= 0 ? 'hi' : 'lo');
        mesVar.textContent  = varMes === null ? 'Sin dato de referencia' : `${varMes >= 0 ? '+' : ''}${varMes.toFixed(2)}%`;
        mesVar.className    = 'pill ' + (varMes === null ? '' : varMes >= 0 ? 'hi' : 'lo');
        fitTables();
    }
    function updateTopArticulos(){
        dc('bubble');
        const filtered = getFilteredTopArt();
        if(!filtered.length){
            document.getElementById('tabla-top').innerHTML =
                '<tr><td colspan="7" class="empty">Selecciona canales de venta y meses para ver el ranking.</td></tr>';
            ['tk-uni','tk-ventas','tk-margen'].forEach(id => {
                document.getElementById(id).textContent = '—';
            });
            fitTables();
            return;
        }
        const artMap = {};
        filtered.forEach(r => {
            if(!artMap[r.Artículo]) artMap[r.Artículo] = {
                Artículo:    r.Artículo,
                Descripción: r.Descripción,
                Fabricante:  r.Fabricante,
                unidades: 0,
                ventas:   0,
                utilidad: 0,
            };
            artMap[r.Artículo].unidades += r.unidades;
            artMap[r.Artículo].ventas   += r.ventas;
            artMap[r.Artículo].utilidad += r.utilidad;
        });
        const top10 = Object.values(artMap)
            .sort((a,b) => b.unidades - a.unidades)
            .slice(0, 10);
        const histFiltrado    = getFilteredHist();
        const totalHistVentas = histFiltrado.reduce((a,r) => a + r.ventas, 0);
        const totUni    = top10.reduce((a,r) => a+r.unidades, 0);
        const totVentas = top10.reduce((a,r) => a+r.ventas,   0);
        document.getElementById('tk-uni').textContent    = fN(totUni);
        document.getElementById('tk-ventas').textContent = fF(totVentas);
        document.getElementById('tk-margen').textContent = fP(totalHistVentas>0 ? totVentas/totalHistVentas : 0);
        const tbody = document.getElementById('tabla-top');
        const rows = top10.map((r, i) => {
            const rank = i + 1;
            const rc   = rank===1?'r1':rank===2?'r2':rank===3?'r3':'';
            const pct  = totalHistVentas>0 ? r.ventas/totalHistVentas : 0;
            return `<tr>
              <td data-label="#" style="text-align:right"><span class="badge-rank ${rc}">${rank}</span></td>
              <td class="art-code" data-label="Artículo">${r.Artículo}</td>
              <td class="art-desc" data-label="Descripción">${r.Descripción}</td>
              <td class="art-fab" data-label="Fabricante">${r.Fabricante}</td>
              <td class="r" data-label="Unidades"><b>${fN(r.unidades)}</b></td>
              <td class="r" data-label="Venta $">${fF(r.ventas)}</td>
              <td class="r" data-label="% Venta"><span class="pill hi">${fP(pct)}</span></td>
            </tr>`;
        }).join('');
        tbody.innerHTML = rows;
        const maxUni = Math.max(...top10.map(r => r.unidades)) || 1;
        const minUni = Math.min(...top10.map(r => r.unidades)) || 0;
        const minR = 10, maxR = 44;
        const rScale = v => {
            const t       = (v - minUni) / (maxUni - minUni || 1);
            const areaMin = minR * minR, areaMax = maxR * maxR;
            return Math.sqrt(areaMin + t * (areaMax - areaMin));
        };
        const bubbleDatasets = top10.map((r, i) => {
            const margen = r.ventas > 0 ? +(r.utilidad / r.ventas * 100).toFixed(2) : 0;
            return {
                label: r.Artículo,
                data: [{ x: r.ventas, y: margen, r: rScale(r.unidades) }],
                backgroundColor: PALETTE_BUBBLE[i % PALETTE_BUBBLE.length] + 'bb',
                borderColor:     PALETTE_BUBBLE[i % PALETTE_BUBBLE.length],
                borderWidth: 1.5,
                _meta: { desc: r.Descripción, fab: r.Fabricante, uni: r.unidades },
            };
        });
        charts['bubble'] = new Chart(document.getElementById('chart-bubble'), {
            type: 'bubble',
            data: { datasets: bubbleDatasets },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                interaction: { mode: 'nearest', intersect: true },
                plugins: {
                    legend: {
                        display: true,
                        position: 'bottom',
                        onClick: null,
                        labels: {
                            boxWidth: 9, boxHeight: 9, usePointStyle: true, pointStyle: 'circle',
                            padding: 12, color: '#616161',
                            font: { family: "'Segoe UI', sans-serif", size: 9.5, weight: '600' }
                        }
                    },
                    datalabels: { display: false },
                    tooltip: {
                        ...PREMIUM_TOOLTIP_OPTS,
                        callbacks: {
                            title: ctx => ctx[0].dataset.label,
                            label: ctx => {
                                const ds  = ctx.dataset;
                                const pt  = ctx.raw;
                                return [
                                    ` ${ds._meta.desc}`,
                                    ` Fabricante: ${ds._meta.fab}`,
                                    ` Unidades: ${fN(ds._meta.uni)}`,
                                    ` Ventas: ${fF(pt.x)}`,
                                    ` Margen: ${pt.y.toFixed(2)}%`,
                                ];
                            }
                        }
                    }
                },
                scales: {
                    x: {
                        title: { display: true, text: 'Ventas $ c/Desc', color: '#858585', font: { size: 10, weight: '600' } },
                        grid: { color: '#f4f4f4' },
                        ticks: { font: { size: 9 }, color: '#858585', callback: v => fM(v) }
                    },
                    y: {
                        title: { display: true, text: 'Margen %', color: '#858585', font: { size: 10, weight: '600' } },
                        grid: { color: '#f4f4f4' },
                        ticks: { font: { size: 9 }, color: '#858585', callback: v => v.toFixed(1)+'%' }
                    }
                }
            }
        });
        fitTables();
    }
    window.toggleLineaCat = function(rid){
        const rows = document.querySelectorAll('.'+rid+'-child');
        if(!rows.length) return;
        const show = rows[0].style.display === 'none';
        rows.forEach(r => { r.style.display = show ? 'table-row' : 'none'; });
        const icon = document.getElementById(rid+'-icon');
        if(icon) icon.textContent = show ? '▾' : '▸';
        fitTables();
    };
    window.toggleFabQuartil = function(rid){
        const rows = document.querySelectorAll('.'+rid+'-child');
        if(!rows.length) return;
        const show = rows[0].style.display === 'none';
        rows.forEach(r => { r.style.display = show ? 'table-row' : 'none'; });
        const icon = document.getElementById(rid+'-icon');
        if(icon) icon.textContent = show ? '▾' : '▸';
        fitTables();
    };
    function updateLineasCategoria(){
        dc('treemap');
        const tbody    = document.getElementById('tabla-lineascategoria');
        const filtered = getFilteredLineasCat();
        if(!filtered.length){
            tbody.innerHTML = '<tr><td colspan="6" class="empty">Selecciona canales de venta y meses para ver el análisis.</td></tr>';
            document.getElementById('tk2-linea-nombre').textContent = '—';
            document.getElementById('tk2-linea-sub').textContent    = '—';
            document.getElementById('tk2-cat-nombre').textContent   = '—';
            document.getElementById('tk2-cat-sub').textContent      = '—';
            document.getElementById('tk2-cat-margen').textContent   = '—';
            fitTables();
            return;
        }
        const key2 = {};
        const lineaTotales = {};
        filtered.forEach(r => {
            const k = r.Línea + '|||' + r.Categoría;
            if(!key2[k]) key2[k] = {Línea:r.Línea, Categoría:r.Categoría, ventas:0, utilidad:0, unidades:0};
            key2[k].ventas    += r.ventas;
            key2[k].utilidad  += r.utilidad;
            key2[k].unidades  += r.unidades;
            lineaTotales[r.Línea] = (lineaTotales[r.Línea]||0) + r.ventas;
        });
        const allLeaves = Object.values(key2);
        const totalSeleccion = filtered.reduce((a,r) => a+r.ventas, 0);
        const lineaLiderEntry = Object.entries(lineaTotales).sort((a,b) => b[1]-a[1])[0];
        if(lineaLiderEntry){
            const [lineaLiderNombre, lineaLiderVentas] = lineaLiderEntry;
            const lineaLiderPct = totalSeleccion>0 ? lineaLiderVentas/totalSeleccion : 0;
            document.getElementById('tk2-linea-nombre').textContent = lineaLiderNombre;
            document.getElementById('tk2-linea-sub').textContent    = `${fF(lineaLiderVentas)} · ${fP(lineaLiderPct)} de participación`;
        }
        const catLiderEntry = allLeaves.slice().sort((a,b) => b.ventas - a.ventas)[0];
        if(catLiderEntry){
            const catLiderPct = totalSeleccion>0 ? catLiderEntry.ventas/totalSeleccion : 0;
            const catLiderMg  = catLiderEntry.ventas>0 ? catLiderEntry.utilidad/catLiderEntry.ventas : 0;
            document.getElementById('tk2-cat-nombre').textContent = catLiderEntry.Categoría;
            document.getElementById('tk2-cat-sub').textContent    = `${catLiderEntry.Línea} · ${fF(catLiderEntry.ventas)} · ${fP(catLiderPct)} de participación`;
            document.getElementById('tk2-cat-margen').textContent = fP(catLiderMg);
        }
        const lineaMap = {};
        allLeaves.forEach(r => {
            if(!lineaMap[r.Línea]) lineaMap[r.Línea] = {Línea:r.Línea, ventas:0, utilidad:0, unidades:0, cats:[]};
            lineaMap[r.Línea].ventas   += r.ventas;
            lineaMap[r.Línea].utilidad += r.utilidad;
            lineaMap[r.Línea].unidades += r.unidades;
            lineaMap[r.Línea].cats.push(r);
        });
        const lineasSorted = Object.values(lineaMap).sort((a,b) => b.ventas-a.ventas);
        let html = '';
        lineasSorted.forEach((l, i) => {
            const mg  = l.ventas>0 ? l.utilidad/l.ventas : 0;
            const rid = 'lc'+i;
            html += `<tr class="lc-parent" onclick="toggleLineaCat('${rid}')">
                <td data-label="" style="text-align:right"><span class="lc-icon" id="${rid}-icon">▸</span></td>
                <td data-label="Línea"><b>${l.Línea}</b></td>
                <td class="r" data-label="Unidades">${fN(l.unidades)}</td>
                <td class="r" data-label="Ventas $"><b>${fF(l.ventas)}</b></td>
                <td class="r" data-label="Utilidad">${fF(l.utilidad)}</td>
                <td class="r" data-label="Margen"><span class="pill ${mg>=.50?'hi':'mi'}">${fP(mg)}</span></td>
            </tr>`;
            const catsSorted = l.cats.slice().sort((a,b) => b.ventas-a.ventas);
            catsSorted.forEach(c => {
                const cmg = c.ventas>0 ? c.utilidad/c.ventas : 0;
                html += `<tr class="lc-child ${rid}-child" style="display:none">
                    <td data-label=""></td>
                    <td data-label="Categoría">${c.Categoría}</td>
                    <td class="r" data-label="Unidades">${fN(c.unidades)}</td>
                    <td class="r" data-label="Ventas $">${fF(c.ventas)}</td>
                    <td class="r" data-label="Utilidad">${fF(c.utilidad)}</td>
                    <td class="r" data-label="Margen"><span class="pill ${cmg>=.50?'hi':'mi'}">${fP(cmg)}</span></td>
                </tr>`;
            });
        });
        tbody.innerHTML = html;
        const treemapLeaves = [];
        Object.keys(lineaTotales).forEach(lineaName => {
            const topCats = allLeaves
                .filter(r => r.Línea === lineaName)
                .sort((a,b) => b.ventas - a.ventas)
                .slice(0, 3);
            treemapLeaves.push(...topCats);
        });
        const treemapLineaTotales = {};
        treemapLeaves.forEach(r => { treemapLineaTotales[r.Línea] = (treemapLineaTotales[r.Línea]||0) + r.ventas; });
        const lineaVals = Object.values(treemapLineaTotales);
        const leafVals  = treemapLeaves.map(r => r.ventas);
        const lineaMin  = Math.min(...lineaVals), lineaMax = Math.max(...lineaVals);
        const leafMin   = Math.min(...leafVals),  leafMax  = Math.max(...leafVals);
        const headerValueMap = {};
        Object.keys(treemapLineaTotales).forEach(lineaName => {
            headerValueMap[Math.round(treemapLineaTotales[lineaName])] = lineaName;
        });
        const numCeldas = treemapLeaves.length;
        const alturaTreemap = Math.max(360, Math.min(900, 360 + Math.max(0, numCeldas - 40) * 3));
        document.getElementById('lc-chart-wrap').style.height = alturaTreemap + 'px';
        charts['treemap'] = new Chart(document.getElementById('chart-treemap'), {
            type: 'treemap',
            data: {
                datasets: [{
                    label: 'Ventas c/Desc',
                    tree: treemapLeaves,
                    key: 'ventas',
                    groups: ['Línea','Categoría'],
                    spacing: 1,
                    borderWidth: 1.5,
                    borderColor: '#ffffff',
                    displayMode: 'headerBoxes',
                    captions: {
                        display: true,
                        color: '#ffffff',
                        font: { weight: '700', size: 10 },
                        formatter: ctx => ctx.raw.g
                    },
                    labels: {
                        display: ctx => ctx.type === 'data' && ctx.raw.l > 0 && ctx.raw.w > 46 && ctx.raw.h > 18,
                        color: ctx => {
                            if(ctx.type !== 'data' || ctx.raw.l === 0) return '#ffffff';
                            const v = ctx.raw.v || 0;
                            const t = leafMax > leafMin ? (v-leafMin)/(leafMax-leafMin) : 1;
                            return t < 0.45 ? '#262626' : '#ffffff';
                        },
                        font: { size: 8.5, weight: '600' },
                        overflow: 'hidden',
                        formatter: ctx => {
                            if(ctx.type !== 'data' || ctx.raw.l === 0) return '';
                            return ctx.raw.g;
                        }
                    },
                    backgroundColor: ctx => {
                        if(ctx.type !== 'data') return 'transparent';
                        const v = ctx.raw.v || 0;
                        if(ctx.raw.l === 0){
                            const t = lineaMax > lineaMin ? (v-lineaMin)/(lineaMax-lineaMin) : 1;
                            return salesColor(0.55 + t*0.45);
                        }
                        const t = leafMax > leafMin ? (v-leafMin)/(leafMax-leafMin) : 1;
                        return salesColor(t);
                    },
                }]
            },
            options: {
                responsive: true,
                maintainAspectRatio: false,
                plugins: {
                    datalabels: { display: false},
                    legend: { display: false },
                    tooltip: {
                        ...PREMIUM_TOOLTIP_OPTS,
                        callbacks: {
                            title: ctx => {
                                const raw = ctx[0] && ctx[0].raw;
                                if(!raw) return '';
                                if(raw.l === 0){
                                    const nombre = raw.g || headerValueMap[Math.round(raw.v || raw.s || 0)] || '';
                                    return nombre ? `Línea: ${nombre}` : 'Línea';
                                }
                                const orig = raw._data;
                                if(orig && orig.Línea && orig.Categoría) return `${orig.Línea} · ${orig.Categoría}`;
                                return raw.g || 'Categoría';
                            },
                            label: ctx => ' Ventas: ' + fF(ctx.raw.v || 0)
                        }
                    }
                }
            }
        });
        fitTables();
    }
    const QUART_LABELS = {1:'Cuartil 1 · Mayor venta', 2:'Cuartil 2', 3:'Cuartil 3', 4:'Cuartil 4 · Menor venta'};
    const QUART_COLORS = {
        1: {bg:'#3D3D3D', fg:'#ffffff', tint:'#3D3D3D14'},
        2: {bg:'#6B6B6B', fg:'#ffffff', tint:'#6B6B6B12'},
        3: {bg:'#B8B8B8', fg:'#262626', tint:'#B8B8B820'},
        4: {bg:'#E8E8E8', fg:'#5A5A5A', tint:'#E8E8E840'},
    };
    function updateFabricantes(){
        dc('fabricantes');
        const tbody    = document.getElementById('tabla-fabricantes');
        const filtered = getFilteredFabricantes();
        if(!filtered.length){
            tbody.innerHTML = '<tr><td colspan="5" class="empty">Selecciona canales de venta y meses para ver los fabricantes.</td></tr>';
            ['tk3-ventas','tk3-margen','tk3-pct'].forEach(id => { document.getElementById(id).textContent = '—'; });
            document.getElementById('tk3-pct-sub').textContent = 'Necesarios para alcanzar el 50% de ventas';
            fitTables();
            return;
        }
        const fabMap = {};
        let totalSeleccion = 0;
        filtered.forEach(r => {
            if(!fabMap[r.Fabricante]) fabMap[r.Fabricante] = {Fabricante:r.Fabricante, ventas:0, utilidad:0, unidades:0};
            fabMap[r.Fabricante].ventas   += r.ventas;
            fabMap[r.Fabricante].utilidad += r.utilidad;
            fabMap[r.Fabricante].unidades += r.unidades;
            totalSeleccion += r.ventas;
        });
        const sorted = Object.values(fabMap).sort((a,b) => b.ventas - a.ventas);
        let cum = 0;
        sorted.forEach(f => {
            const start = totalSeleccion>0 ? cum/totalSeleccion : 0;
            cum += f.ventas;
            f.pct    = totalSeleccion>0 ? f.ventas/totalSeleccion : 0;
            f.cumPct = totalSeleccion>0 ? cum/totalSeleccion : 0;
            if(start < 0.25)      f.cuartil = 1;
            else if(start < 0.50) f.cuartil = 2;
            else if(start < 0.75) f.cuartil = 3;
            else                  f.cuartil = 4;
        });
        const paretoList = [];
        let cumPareto = 0;
        for(const f of sorted){
            paretoList.push(f);
            cumPareto += f.ventas;
            if(totalSeleccion>0 && cumPareto/totalSeleccion >= 0.5) break;
        }
        const totVentasPareto = paretoList.reduce((a,r) => a+r.ventas,   0);
        const totUtilPareto   = paretoList.reduce((a,r) => a+r.utilidad, 0);
        document.getElementById('tk3-ventas').textContent = fF(totVentasPareto);
        document.getElementById('tk3-margen').textContent = fP(totVentasPareto>0 ? totUtilPareto/totVentasPareto : 0);
        document.getElementById('tk3-pct').textContent    = `${paretoList.length} de ${sorted.length}`;
        document.getElementById('tk3-pct-sub').textContent = `Cubren el ${fP(totalSeleccion>0 ? totVentasPareto/totalSeleccion : 0)} de las ventas`;
        const groups = {1:[], 2:[], 3:[], 4:[]};
        sorted.forEach(f => groups[f.cuartil].push(f));
        const maxVentas = sorted.length ? sorted[0].ventas : 0;
        let html = '';
        let cumBeforeGroup = 0;
        [1,2,3,4].forEach(q => {
            const items = groups[q];
            if(!items.length) return;
            const qVentas = items.reduce((a,r) => a+r.ventas,   0);
            const qUtil   = items.reduce((a,r) => a+r.utilidad, 0);
            const qMg     = qVentas>0 ? qUtil/qVentas : 0;
            const qPct    = totalSeleccion>0 ? qVentas/totalSeleccion : 0;
            const rid     = 'fq'+q;
            const qc      = QUART_COLORS[q];
            const expanded = q === 1;
            const startPct = totalSeleccion>0 ? cumBeforeGroup/totalSeleccion : 0;
            cumBeforeGroup += qVentas;
            const endPct   = totalSeleccion>0 ? cumBeforeGroup/totalSeleccion : 0;
            html += `<tr class="fq-parent" onclick="toggleFabQuartil('${rid}')" style="background:${qc.tint}">
                <td data-label="" style="text-align:right"><span class="lc-icon" id="${rid}-icon">${expanded?'▾':'▸'}</span></td>
                <td data-label="Cuartil"><span class="fq-badge" style="background:${qc.bg};color:${qc.fg}">${QUART_LABELS[q]}</span>
                    <span style="color:#858585;font-size:.68rem;margin-left:6px;display:block;text-align:left;margin-top:4px">${items.length} fabricante${items.length!==1?'s':''} · ${fP(startPct)}–${fP(endPct)} acumulado</span></td>
                <td class="r" data-label="Ventas $"><b>${fF(qVentas)}</b></td>
                <td class="r" data-label="Margen"><span class="pill ${qMg>=.50?'hi':'mi'}">${fP(qMg)}</span></td>
                <td class="r" data-label="% Participación"><b>${fP(qPct)}</b></td>
            </tr>`;
            items.forEach((f, idx) => {
                const mg  = f.ventas>0 ? f.utilidad/f.ventas : 0;
                const barW = maxVentas>0 ? (f.ventas/maxVentas*100).toFixed(1) : 0;
                html += `<tr class="lc-child ${rid}-child" style="display:${expanded?'table-row':'none'}">
                    <td class="fab-rank-num" data-label="#" style="text-align:right">${idx+1}</td>
                    <td data-label="Fabricante">${f.Fabricante}</td>
                    <td class="r fab-bar-cell" data-label="Ventas $"><span class="fab-bar-bg" style="width:${barW}%;background:${qc.bg}22"></span><b>${fF(f.ventas)}</b></td>
                    <td class="r" data-label="Margen"><span class="pill ${mg>=.50?'hi':'mi'}">${fP(mg)}</span></td>
                    <td class="r" data-label="% Participación">${fP(f.pct)}</td>
                </tr>`;
            });
        });
        tbody.innerHTML = html;
        document.getElementById('fab-chart-wrap').style.height = Math.max(220, paretoList.length*34 + 60) + 'px';
        const maxVal = Math.max(...paretoList.map(f => f.ventas)) || 0;
        charts['fabricantes'] = new Chart(document.getElementById('chart-fabricantes'), {
            type: 'bar',
            data: {
                labels: paretoList.map(f => f.Fabricante),
                datasets: [{
                    data: paretoList.map(f => f.ventas),
                    backgroundColor: paretoList.map((f,i) => QUART_COLORS[f.cuartil].bg),
                    borderRadius: 5,
                    barThickness: 16
                }]
            },
            options: {
                indexAxis: 'y', responsive: true, maintainAspectRatio: false,
                layout: { padding: { right: 34 } },
                plugins: {
                    legend: { display: false },
                    tooltip: {
                        ...PREMIUM_TOOLTIP_OPTS,
                        callbacks: {
                            label: ctx => {
                                const f  = paretoList[ctx.dataIndex];
                                const mg = f.ventas>0 ? f.utilidad/f.ventas : 0;
                                return [' Ventas: ' + fF(f.ventas), ' Margen: ' + fP(mg), ' % Participación: ' + fP(f.pct)];
                            }
                        }
                    },
                    datalabels: {
                        display: true, anchor: 'end', align: 'end',
                        color: '#4d4d4d', font: { weight: '600', size: 9.5 },
                        formatter: (value, ctx) => fP(paretoList[ctx.dataIndex].pct)
                    }
                },
                scales: {
                    x: { grid:{display:false}, border:{display:false}, suggestedMax:maxVal*1.15, ticks:{font:{size:9},color:'#858585',callback:v=>fM(v)} },
                    y: { grid:{display:false}, border:{display:false}, ticks:{font:{size:10,weight:'600'},color:'#212121'} }
                }
            }
        });
        fitTables();
    }
    function updateResumenAnterior(slug){
        const info = RESUMENES_ANT[slug];
        if(!info) return;
        const data  = info.data.filter(r => active.has(r.NombreSucursal));
        const dates = [...new Set(data.map(r => r.FechaStr))].sort();
        const byDate = aggByDate(data);
        const tbody = document.getElementById('tabla-resumen-' + slug);
        if(!tbody) return;
        const valid = dates.filter(d => byDate[d]);
        if(!valid.length){
            tbody.innerHTML = '<tr><td colspan="7" class="empty">Selecciona canales de venta para ver el detalle.</td></tr>';
            dc('resumen-'+slug+'-ventas');
            dc('resumen-'+slug+'-tickets');
            fitTables();
            return;
        }
        let tUn=0,tV=0,tU=0,tTk=0;
        const rows = valid.map(d => {
            const r  = byDate[d];
            const mg = r.ventas>0 ? r.utilidad/r.ventas : 0;
            const pill = mg>=.50 ? 'hi' : 'mi';
            const dn = DAYS[new Date(d+'T12:00:00').getDay()];
            const mAbr = MESES_ABR[parseInt(d.slice(5,7),10)-1];
            tUn+=r.unidades; tV+=r.ventas; tU+=r.utilidad; tTk+=r.tickets;
            return `<tr><td class="date" data-label="Fecha">${d.slice(8)} ${mAbr}</td><td class="dayname" data-label="Día">${dn}</td><td class="r" data-label="Unidades">${fN(r.unidades)}</td><td class="r" data-label="Ventas $"><b>${fF(r.ventas)}</b></td><td class="r" data-label="Utilidad">${fF(r.utilidad)}</td><td class="r" data-label="Margen"><span class="pill ${pill}">${fP(mg)}</span></td><td class="r" data-label="Tickets"><b>${fN(r.tickets)}</b></td></tr>`;
        }).join('');
        const totMg = tV>0 ? tU/tV : 0;
        tbody.innerHTML = rows + `<tr class="total-row"><td data-label="" colspan="2"><b>TOTAL ${info.label.toUpperCase()} ${info.anio}</b></td><td class="r" data-label="Unidades">${fN(tUn)}</td><td class="r" data-label="Ventas $">${fF(tV)}</td><td class="r" data-label="Utilidad">${fF(tU)}</td><td class="r" data-label="Margen"><span class="pill ${totMg>=.50?'hi':'mi'}">${fP(totMg)}</span></td><td class="r" data-label="Tickets">${fN(tTk)}</td></tr>`;
        const activeSucs = getActiveSucs();
        const multi      = activeSucs.length > 1;
        updateChartGeneric('resumen-'+slug+'-ventas',  'ventas',  data, valid, byDate, activeSucs, multi);
        updateChartGeneric('resumen-'+slug+'-tickets', 'tickets', data, valid, byDate, activeSucs, multi);
        fitTables();
    }
    buildButtons();
    buildMesButtons();
    refreshMesesDisponibilidad();
    refreshSucursalesDisponibilidad();
    document.getElementById('filter-bar-mes').style.display = 'none';
    document.getElementById('filter-bar-suc').style.display = 'none';
    updatePronostico();
    setTimeout(fitTables, 100);
});
</script>
</body>
</html>"""
    html = html_template.replace("__FECHA_VALOR__",   fecha_reporte)
    html = html.replace("__FECHA_INFO__",             fecha_info)
    html = html.replace("__MES_HEADER__",             mes_header)
    html = html.replace("__DATA_JSON__",              data_json)
    html = html.replace("__LINEA_JSON__",             linea_json)
    html = html.replace("__HISTORICO_JSON__",         historico_json)
    html = html.replace("__TOP_ART_JSON__",           top_art_json)
    html = html.replace("__LINEAS_CAT_JSON__",        lineas_cat_json)
    html = html.replace("__FABRICANTES_JSON__",       fabricantes_json)
    html = html.replace("__PRONOSTICO_JSON__",        pronostico_json)
    html = html.replace("__PERIODOS_JSON__",          periodos_json)
    html = html.replace("__SUCURSALES_JSON__",        sucursales_json)
    html = html.replace("__CURRENT_PERIOD_JSON__",    current_period_json)
    html = html.replace("__TICKETS_CONS_JSON__",      tickets_cons_json)
    html = html.replace("__TABS_NAV_EXTRA__",         tabs_nav_extra)
    html = html.replace("__TABS_CONTENT_EXTRA__",     tabs_content_extra)
    html = html.replace("__RESUMENES_ANT_JSON__",     resumenes_ant_json)
    html = html.replace("__LOGO_B64__",               logo_b64)
    return html
def main():
    try:
        print(f"Leyendo: {EXCEL_PATH}")
        xl  = pd.read_excel(EXCEL_PATH, sheet_name=None)
        vmc = xl["VentasMesCurso"].copy()
        vm  = xl["VentasMensuales"].copy()
        tkt = xl["TicketsMensuales"].copy()
        dim = xl["TablasDimensión"].copy()
        objetivos_df = xl.get("ObjetivosVentas")
        if objetivos_df is None:
            print("⚠️  No se encontró la hoja 'ObjetivosVentas'; el pronóstico y el orden del segmentador "
                  "se calcularán sin fechas de apertura (se infieren con la primera venta del mes).")
            objetivos_df = pd.DataFrame(columns=["Sucursal", "FechaApertura"])
        else:
            objetivos_df = objetivos_df.copy()
        suc = dim[["Clave sucursal","Nombre de sucursal"]].dropna().drop_duplicates()
        suc.columns = ["ClaveSucursal","NombreSucursal"]
        suc["ClaveSucursal"] = pd.to_numeric(suc["ClaveSucursal"], errors="coerce").fillna(0).astype(int)
        lista_sucursales = ordenar_sucursales_por_apertura(suc, objetivos_df)
        art     = dim[["Artículo","Línea"]].dropna().drop_duplicates("Artículo")
        art_dim = dim[["Artículo","Línea","Categoría","Descripción","Fabricante"]].dropna(subset=["Artículo"]).drop_duplicates("Artículo")
        print("Procesando datos...")
        vmc_actual, meses_anteriores = separar_mes_actual_anterior(vmc, FECHA_BASE)
        agg = procesar_mes_curso(vmc_actual, suc, BOL_EXCLUIR)
        resumenes_anteriores = []
        for info in meses_anteriores:
            agg_ant = procesar_mes_curso(info["df"], suc, BOL_EXCLUIR)
            resumenes_anteriores.append({
                "label": info["label"], "anio": info["anio"], "slug": info["slug"], "agg": agg_ant
            })
        linea_agg       = procesar_lineas(vm, art, suc, BOL_EXCLUIR)
        historico_agg, tickets_consolidados = procesar_historico(vm, tkt, suc, BOL_EXCLUIR)
        top_art_agg     = procesar_top_articulos(vm, art_dim, suc, BOL_EXCLUIR)
        lineas_cat_agg  = procesar_lineas_categoria(vm, art_dim, suc, BOL_EXCLUIR)
        fabricantes_agg = procesar_fabricantes(vm, art_dim, suc, BOL_EXCLUIR)
        pronostico_agg  = procesar_pronostico(agg, objetivos_df, suc, FECHA_BASE, es_cierre=ES_CIERRE_MES)
        print("Generando HTML final...")
        fecha_reporte, fecha_info, mes_header = formatear_fechas(FECHA_BASE, es_cierre=ES_CIERRE_MES)
        current_period_label = periodo_label_actual(FECHA_BASE)
        html = generar_html(agg, linea_agg, historico_agg, top_art_agg, lineas_cat_agg, fabricantes_agg,
                            pronostico_agg, lista_sucursales, fecha_reporte, fecha_info, mes_header,
                            current_period_label, tickets_consolidados, resumenes_anteriores)
        Path(OUTPUT_PATH).write_text(html, encoding="utf-8")
        print(f"✅ Dashboard generado exitosamente en: {OUTPUT_PATH}")
    except Exception as e:
        import traceback
        print(f"❌ Error crítico durante la ejecución: {e}")
        traceback.print_exc()
if __name__ == "__main__":
    main()