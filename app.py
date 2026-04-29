import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
import plotly.express as px
from datetime import date

# ============================================================
# CONFIG
# ============================================================
st.set_page_config(page_title="PRE-S&OP | 12M GAP Board", layout="wide")

# =========================
# Columnas (con Periodo para 12 meses)
# =========================
COL_PERIOD  = "Periodo"  # YYYY-MM
HIER = ["Compañía", "País", "Categoría", "Marca", "Marca BU", "SKU", "Descripción"]

COL_FORECAST = "Plan de Demanda Ajustada (V2) - Cajas"
COL_SUPPLY   = "Plan de Abastecimiento - Cajas"
COL_GAP      = "GAP Demanda y Abastecimiento Cajas"
COL_FR_PROJ  = "Fill Rate Proyectado"
COL_FR_OBJ   = "Fill Rate Objetivo"
COL_MARGIN   = "Precio Unitario $USD"
COL_COST     = "Costo Unitario $USD"
COL_IMPACT   = "Impacto Economico $USD"
COL_SEM      = "Semáforo"

COL_RESTR = "Restricción Principal"
COL_ACTION = "Acción Propuesta"
COL_ESC = "Escalar a S&OP ejecutivo"

EDITABLE = [COL_FORECAST, COL_SUPPLY, COL_RESTR, COL_ACTION, COL_ESC]
SEM_COLORS = {"ROJO":"#D72638", "AMARILLO":"#F6C343", "VERDE":"#1F9D55"}

# =========================
# UI helper
# =========================
st.markdown("""
<style>
.stApp { background:#F7F9FC; }
.header{
  background:#071B3A;color:white;border-radius:16px;padding:16px 18px;
  box-shadow:0 10px 22px rgba(7,27,58,0.18)
}
.header h2{margin:0;font-weight:900}
.header .sub{opacity:.88;margin-top:6px}
.kpi{background:white;border:1px solid rgba(0,0,0,.08);border-radius:14px;padding:10px 12px}
.kpi .t{opacity:.6;font-weight:700;font-size:.8rem}
.kpi .v{font-weight:900;font-size:1.4rem}
</style>
""", unsafe_allow_html=True)

def kpi(title, value):
    st.markdown(f"""<div class="kpi"><div class="t">{title}</div><div class="v">{value}</div></div>""",
                unsafe_allow_html=True)

def to_excel_bytes(sheets: dict[str, pd.DataFrame]) -> bytes:
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="openpyxl") as writer:
        for name, df_ in sheets.items():
            df_.to_excel(writer, index=False, sheet_name=name[:31])
    return bio.getvalue()

def safe_num(s):
    return pd.to_numeric(s, errors="coerce").fillna(0)

def ensure_cols(df: pd.DataFrame, cols: list[str], default=np.nan):
    df = df.copy()
    for c in cols:
        if c not in df.columns:
            df[c] = default
    return df

def normalize_period(df: pd.DataFrame) -> pd.DataFrame:
    """
    Periodo debe ser YYYY-MM.
    Si viene fecha, lo convertimos.
    Si no existe, lo marcamos (SinPeriodo).
    """
    df = df.copy()
    if COL_PERIOD not in df.columns:
        df[COL_PERIOD] = "(SinPeriodo)"
        return df

    s = df[COL_PERIOD]

    # intentar parse como fecha
    dt = pd.to_datetime(s, errors="coerce")
    if dt.notna().any():
        df[COL_PERIOD] = dt.dt.to_period("M").astype(str)
        df[COL_PERIOD] = df[COL_PERIOD].fillna("(SinPeriodo)")
        return df

    # string cleanup
    s2 = s.astype(str).str.strip()
    s2 = s2.str.replace("/", "-", regex=False)
    # 202501 -> 2025-01
    s2 = s2.str.replace(r"^(\d{4})(\d{2})$", r"\1-\2", regex=True)
    df[COL_PERIOD] = s2.replace({"nan":"(SinPeriodo)", "None":"(SinPeriodo)", "":"(SinPeriodo)"})
    return df

def dedup_safe(df: pd.DataFrame, key_cols: list[str], sum_cols: list[str]) -> pd.DataFrame:
    """
    Deduplicación segura: si hay llaves repetidas, agrega sumando numéricos y manteniendo first en texto.
    """
    df = df.copy()
    if df.duplicated(subset=key_cols).any():
        agg = {c: "sum" for c in sum_cols if c in df.columns}
        for c in df.columns:
            if c not in key_cols and c not in agg:
                agg[c] = "first"
        df = df.groupby(key_cols, dropna=False, as_index=False).agg(agg)
    return df

# =========================
# FIX CLAVE: dtype string para Restricción/Acción/Escalar (evita StreamlitAPIException)
# =========================
def compute_all(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    # asegurar columnas
    df = ensure_cols(df, HIER, default="")
    df = ensure_cols(df, [COL_FR_OBJ, COL_MARGIN, COL_COST], default=0)
    df = ensure_cols(df, [COL_RESTR, COL_ACTION, COL_ESC], default="")

    # numéricos base
    df[COL_FORECAST] = safe_num(df[COL_FORECAST])
    df[COL_SUPPLY]   = safe_num(df[COL_SUPPLY])
    df[COL_MARGIN]   = safe_num(df[COL_MARGIN])
    df[COL_COST]     = safe_num(df[COL_COST])
    df[COL_FR_OBJ]   = pd.to_numeric(df[COL_FR_OBJ], errors="coerce")

    # defaults jerarquía
    df["Compañía"] = df["Compañía"].replace("", np.nan).fillna("PDC")

    # FORZAR TEXTO (clave)
    df[COL_RESTR] = df[COL_RESTR].fillna("").astype(str)
    df[COL_ACTION] = df[COL_ACTION].fillna("").astype(str)

    df[COL_ESC] = df[COL_ESC].fillna("NO").astype(str).str.upper().str.strip()
    df[COL_ESC] = df[COL_ESC].replace({"SI":"SÍ","SÍ":"SÍ","NO":"NO"})
    df.loc[~df[COL_ESC].isin(["SÍ","NO"]), COL_ESC] = "NO"

    # cálculos
    df[COL_GAP] = df[COL_FORECAST] - df[COL_SUPPLY]
    df[COL_FR_PROJ] = np.where(df[COL_FORECAST] > 0, df[COL_SUPPLY] / df[COL_FORECAST], np.nan)

    if df[COL_MARGIN].sum() > 0:
        df[COL_IMPACT] = df[COL_GAP].abs() * df[COL_MARGIN]
    elif df[COL_COST].sum() > 0:
        df[COL_IMPACT] = df[COL_GAP].abs() * df[COL_COST]
    else:
        df[COL_IMPACT] = df[COL_GAP].abs()

    # semáforo
    df[COL_SEM] = "VERDE"
    df.loc[df[COL_GAP].abs() > 0, COL_SEM] = "AMARILLO"
    has_obj = df[COL_FR_OBJ].notna() & (df[COL_FR_OBJ] > 0)
    df.loc[has_obj & df[COL_FR_PROJ].notna() & (df[COL_FR_PROJ] < df[COL_FR_OBJ]), COL_SEM] = "ROJO"
    df.loc[df[COL_GAP].abs() == 0, COL_SEM] = "VERDE"

    return df

def consolidate(dem: pd.DataFrame, sup: pd.DataFrame) -> pd.DataFrame:
    dem = dem.copy()
    sup = sup.copy()

    # asegurar columnas base
    dem = ensure_cols(dem, [COL_PERIOD] + HIER + [COL_FORECAST, COL_MARGIN, COL_COST, COL_FR_OBJ], default=np.nan)
    sup = ensure_cols(sup, [COL_PERIOD] + HIER + [COL_SUPPLY, COL_RESTR, COL_ACTION, COL_ESC], default=np.nan)

    dem = normalize_period(dem)
    sup = normalize_period(sup)

    key_cols = [COL_PERIOD] + HIER

    # dedup antes del merge (evita explosión)
    dem = dedup_safe(dem, key_cols, sum_cols=[COL_FORECAST])
    sup = dedup_safe(sup, key_cols, sum_cols=[COL_SUPPLY])

    merged = dem.merge(
        sup[key_cols + [COL_SUPPLY, COL_RESTR, COL_ACTION, COL_ESC]],
        on=key_cols,
        how="outer"
    )

    merged = compute_all(merged)
    return merged

def template_demand_12m() -> pd.DataFrame:
    cols = [COL_PERIOD] + HIER + [COL_FORECAST, COL_MARGIN, COL_COST, COL_FR_OBJ]
    return pd.DataFrame(columns=cols)

def template_supply_12m() -> pd.DataFrame:
    cols = [COL_PERIOD] + HIER + [COL_SUPPLY, COL_RESTR, COL_ACTION, COL_ESC]
    return pd.DataFrame(columns=cols)

def template_minutes() -> pd.DataFrame:
    cols = ["Fecha", "Ciclo", "Moderador", "Asistentes",
            "Resumen ejecutivo (bullets)", "Decisiones clave",
            "Riesgos / restricciones relevantes", "Acciones y responsables",
            "Escalaciones a S&OP Ejecutivo", "Notas"]
    return pd.DataFrame(columns=cols)

# =========================
# Header
# =========================
st.markdown("""
<div class="header">
  <h2>PRE-S&OP · 12 Meses · Demand vs Supply</h2>
  <div class="sub">Periodo en la llave + deduplicación + worklist editable + KPIs dinámicos en Drill-down</div>
  <div style="height:4px;background:#12B0A6;border-radius:999px;margin-top:10px"></div>
</div>
""", unsafe_allow_html=True)
st.write("")

# =========================
# Sidebar
# =========================
st.sidebar.markdown("## Sesión")
session_date = st.sidebar.date_input("Fecha sesión", value=date.today())
cycle_name = st.sidebar.text_input("Ciclo / Mes", value="PRE-S&OP 2025")
facilitator = st.sidebar.text_input("Moderador", value="")
top_n = st.sidebar.slider("Top N worklist (para editar)", 20, 500, 120, 20)

# =========================
# Tabs
# =========================
tab1, tab2, tab3, tab4 = st.tabs(["Plantillas", "Carga & Semáforo", "Drill-down & Ajustes", "Export"])

with tab1:
    st.markdown("### Plantillas (incluye Periodo para horizonte 12M)")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.download_button("⬇️ Plantilla Demanda 12M (xlsx)",
                           data=to_excel_bytes({"Plan_Demanda": template_demand_12m()}),
                           file_name="Plantilla_Plan_Demanda_12M.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with c2:
        st.download_button("⬇️ Plantilla Abastecimiento 12M (xlsx)",
                           data=to_excel_bytes({"Plan_Abastecimiento": template_supply_12m()}),
                           file_name="Plantilla_Plan_Abastecimiento_12M.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with c3:
        st.download_button("⬇️ Plantilla Minuta (xlsx)",
                           data=to_excel_bytes({"Minuta_PRE_SOP": template_minutes()}),
                           file_name="Plantilla_Minuta_PRE_SOP.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    st.info("Periodo recomendado: YYYY-MM (ej: 2025-01). Si viene fecha, se convierte automáticamente.")

with tab2:
    st.markdown("### Carga (Demanda + Abastecimiento) + Semáforo")
    st.caption("Vista ejecutiva: NO muestra Restricción/Acción/Escalar. Solo KPIs, ranking y tendencia 12M.")

    u1, u2 = st.columns(2)
    with u1:
        f_dem = st.file_uploader("Subir Plan Demanda (xlsx)", type=["xlsx"], key="dem")
    with u2:
        f_sup = st.file_uploader("Subir Plan Abastecimiento (xlsx)", type=["xlsx"], key="sup")
    f_min = st.file_uploader("Subir Minuta (opcional, xlsx)", type=["xlsx"], key="min")

    if not f_dem or not f_sup:
        st.warning("Sube ambos archivos para continuar.")
        st.stop()

    dem = pd.read_excel(f_dem, engine="openpyxl")
    sup = pd.read_excel(f_sup, engine="openpyxl")

    cons = consolidate(dem, sup)
    st.session_state["cons_base"] = cons.copy()

    if f_min:
        try:
            st.session_state["minuta_df"] = pd.read_excel(f_min, engine="openpyxl")
        except Exception:
            st.session_state["minuta_df"] = pd.DataFrame()
    else:
        st.session_state["minuta_df"] = pd.DataFrame()

    total_f = float(cons[COL_FORECAST].sum())
    total_s = float(cons[COL_SUPPLY].sum())
    total_gap = float(cons[COL_GAP].sum())
    fr_w = (total_s/total_f) if total_f>0 else np.nan
    total_imp = float(cons[COL_IMPACT].sum())

    k1,k2,k3,k4,k5 = st.columns(5)
    with k1: kpi("Forecast", f"{total_f:,.0f}")
    with k2: kpi("Abastecimiento", f"{total_s:,.0f}")
    with k3: kpi("GAP", f"{total_gap:,.0f}")
    with k4: kpi("Fill Rate", f"{(fr_w*100):,.1f}%" if pd.notna(fr_w) else "N/A")
    with k5: kpi("Impacto", f"USD {total_imp:,.0f}")

    st.write("")
    st.markdown("#### Tendencia 12 meses (por Periodo)")
    by_m = cons.groupby(COL_PERIOD, dropna=False).agg(
        Forecast=(COL_FORECAST,"sum"),
        Supply=(COL_SUPPLY,"sum"),
        GAP=(COL_GAP,"sum"),
        Impacto=(COL_IMPACT,"sum"),
    ).reset_index()
    by_m["FillRate"] = np.where(by_m["Forecast"]>0, by_m["Supply"]/by_m["Forecast"], np.nan)

    fig = px.line(by_m, x=COL_PERIOD, y=["Forecast","Supply","GAP"], title="Forecast vs Supply vs GAP (por mes)")
    st.plotly_chart(fig, use_container_width=True)

    st.write("")
    st.markdown("#### Ranking críticos (Top 200)")
    view = cons.drop(columns=[COL_RESTR, COL_ACTION, COL_ESC], errors="ignore").copy()
    view["_GAP_ABS"] = view[COL_GAP].abs()
    view = view.sort_values([COL_IMPACT, "_GAP_ABS"], ascending=False).drop(columns=["_GAP_ABS"])
    st.dataframe(view.head(200), use_container_width=True, height=420)

with tab3:
    st.markdown("### Drill-down & Ajustes (worklist estable + KPIs dinámicos)")
    st.caption("Edita SOLO worklist (Top N). Click en 'Aplicar cambios' para recalcular. KPIs se actualizan con filtros y cambios.")

    if "cons_base" not in st.session_state:
        st.warning("Primero carga archivos en 'Carga & Semáforo'.")
        st.stop()

    base = st.session_state.get("cons_edited", st.session_state["cons_base"]).copy()
    base = compute_all(base)

    # filtros arriba
    f1,f2,f3,f4,f5,f6 = st.columns(6)
    with f1:
        p_opts = sorted(base[COL_PERIOD].dropna().unique().tolist())
        sel_p = st.multiselect("Periodo", p_opts, default=p_opts[:12] if len(p_opts)>12 else p_opts)
    with f2:
        c_opts = sorted(base["Compañía"].dropna().unique().tolist())
        sel_c = st.multiselect("Compañía", c_opts, default=c_opts)
    with f3:
        pa_opts = sorted(base["País"].dropna().unique().tolist())
        sel_pa = st.multiselect("País", pa_opts, default=pa_opts)
    with f4:
        ca_opts = sorted(base["Categoría"].dropna().unique().tolist())
        sel_ca = st.multiselect("Categoría", ca_opts, default=ca_opts)
    with f5:
        m_opts = sorted(base["Marca"].dropna().unique().tolist())
        sel_m = st.multiselect("Marca", m_opts, default=m_opts)
    with f6:
        sem_opts = ["ROJO","AMARILLO","VERDE"]
        sel_sem = st.multiselect("Semáforo", sem_opts, default=sem_opts)

    c7,c8,c9 = st.columns([1,1,1])
    with c7:
        only_gap = st.checkbox("Solo GAP ≠ 0", value=True)
    with c8:
        search_sku = st.text_input("Buscar SKU", value="")
    with c9:
        work_mode = st.selectbox("Worklist", ["Top críticos", "Todo filtrado"], index=0)

    df = base.copy()
    if sel_p: df = df[df[COL_PERIOD].isin(sel_p)]
    if sel_c: df = df[df["Compañía"].isin(sel_c)]
    if sel_pa: df = df[df["País"].isin(sel_pa)]
    if sel_ca: df = df[df["Categoría"].isin(sel_ca)]
    if sel_m: df = df[df["Marca"].isin(sel_m)]
    if sel_sem: df = df[df[COL_SEM].isin(sel_sem)]
    if only_gap: df = df[df[COL_GAP].abs() > 0]
    if search_sku.strip():
        df = df[df["SKU"].astype(str).str.contains(search_sku.strip(), case=False, na=False)]

    df["_GAP_ABS"] = df[COL_GAP].abs()
    df = df.sort_values([COL_IMPACT, "_GAP_ABS"], ascending=False).drop(columns=["_GAP_ABS"])
    work = df.head(top_n).copy() if work_mode == "Top críticos" else df.copy()

    # FORZAR dtypes string (evita StreamlitAPIException)
    for c in [COL_RESTR, COL_ACTION, COL_ESC]:
        if c in work.columns:
            work[c] = work[c].fillna("").astype(str)
    if COL_ESC in work.columns:
        work[COL_ESC] = work[COL_ESC].str.upper().str.strip().replace({"SI":"SÍ"})
        work.loc[~work[COL_ESC].isin(["SÍ","NO"]), COL_ESC] = "NO"

    disabled_cols = [c for c in work.columns if c not in EDITABLE]
    col_cfg = {
        COL_FORECAST: st.column_config.NumberColumn(COL_FORECAST, step=1, format="%.0f"),
        COL_SUPPLY:   st.column_config.NumberColumn(COL_SUPPLY, step=1, format="%.0f"),
        COL_RESTR:    st.column_config.TextColumn(COL_RESTR),
        COL_ACTION:   st.column_config.TextColumn(COL_ACTION),
        COL_ESC:      st.column_config.SelectboxColumn(COL_ESC, options=["NO","SÍ"]),
    }

    st.markdown("#### Editor de sesión (worklist)")
    edited = st.data_editor(
        work,
        use_container_width=True,
        height=520,
        disabled=disabled_cols,
        column_config=col_cfg,
        key="editor_worklist"
    )

    if st.button("✅ Aplicar cambios y recalcular", type="primary"):
        edited = compute_all(edited)

        key_cols = [COL_PERIOD] + HIER
        full = base.merge(
            edited[key_cols + EDITABLE],
            on=key_cols,
            how="left",
            suffixes=("", "_new")
        )
        for c in EDITABLE:
            if f"{c}_new" in full.columns:
                full[c] = np.where(full[f"{c}_new"].notna(), full[f"{c}_new"], full[c])
                full = full.drop(columns=[f"{c}_new"])
        full = compute_all(full)

        st.session_state["cons_edited"] = full.copy()
        st.success("Cambios aplicados ✅")
        st.rerun()

    # KPI cards dinámicos (se actualizan con filtros + cambios)
    st.write("")
    st.markdown("#### KPIs (dinámicos según filtros y cambios aplicados)")
    base2 = st.session_state.get("cons_edited", base).copy()
    base2 = compute_all(base2)
    df2 = base2.copy()

    if sel_p: df2 = df2[df2[COL_PERIOD].isin(sel_p)]
    if sel_c: df2 = df2[df2["Compañía"].isin(sel_c)]
    if sel_pa: df2 = df2[df2["País"].isin(sel_pa)]
    if sel_ca: df2 = df2[df2["Categoría"].isin(sel_ca)]
    if sel_m: df2 = df2[df2["Marca"].isin(sel_m)]
    if sel_sem: df2 = df2[df2[COL_SEM].isin(sel_sem)]
    if only_gap: df2 = df2[df2[COL_GAP].abs() > 0]
    if search_sku.strip():
        df2 = df2[df2["SKU"].astype(str).str.contains(search_sku.strip(), case=False, na=False)]

    total_f = float(df2[COL_FORECAST].sum())
    total_s = float(df2[COL_SUPPLY].sum())
    total_gap = float(df2[COL_GAP].sum())
    fr_w = (total_s / total_f) if total_f > 0 else np.nan
    total_imp = float(df2[COL_IMPACT].sum())

    k1, k2, k3, k4, k5 = st.columns(5)
    with k1: kpi("Forecast", f"{total_f:,.0f}")
    with k2: kpi("Abastecimiento", f"{total_s:,.0f}")
    with k3: kpi("GAP", f"{total_gap:,.0f}")
    with k4: kpi("Fill Rate", f"{(fr_w*100):,.1f}%" if pd.notna(fr_w) else "N/A")
    with k5: kpi("Impacto", f"USD {total_imp:,.0f}")

    st.write("")
    st.markdown("#### Tendencia 12 meses (filtrado)")
    by_m = df2.groupby(COL_PERIOD, dropna=False).agg(
        Forecast=(COL_FORECAST,"sum"),
        Supply=(COL_SUPPLY,"sum"),
        GAP=(COL_GAP,"sum"),
        Impacto=(COL_IMPACT,"sum"),
    ).reset_index()
    by_m["FillRate"] = np.where(by_m["Forecast"]>0, by_m["Supply"]/by_m["Forecast"], np.nan)
    st.dataframe(by_m, use_container_width=True, height=260)

with tab4:
    st.markdown("### Export Reporte Consolidado PRE-S&OP (Excel)")

    if "cons_base" not in st.session_state:
        st.warning("Primero carga archivos.")
        st.stop()

    final_df = st.session_state.get("cons_edited", st.session_state["cons_base"]).copy()
    final_df = compute_all(final_df)

    tmp = final_df.copy()
    tmp["_GAP_ABS"] = tmp[COL_GAP].abs()
    ranking = tmp.sort_values([COL_IMPACT, "_GAP_ABS"], ascending=False).drop(columns=["_GAP_ABS"]).head(200)

    summary = pd.DataFrame([{
        "Fecha": str(session_date),
        "Ciclo": cycle_name,
        "Moderador": facilitator,
        "Forecast_total": float(final_df[COL_FORECAST].sum()),
        "Supply_total": float(final_df[COL_SUPPLY].sum()),
        "GAP_total": float(final_df[COL_GAP].sum()),
        "Impacto_total": float(final_df[COL_IMPACT].sum()),
    }])

    minuta_df = st.session_state.get("minuta_df", pd.DataFrame())

    report = to_excel_bytes({
        "01_Resumen": summary,
        "02_Ranking_Criticos": ranking,
        "03_Consolidado": final_df,
        "04_Minuta": minuta_df
    })

    st.download_button(
        "⬇️ Descargar Reporte PRE-S&OP (xlsx)",
        data=report,
        file_name=f"Reporte_PRE_SOP_12M_{cycle_name.replace(' ','_')}_{session_date}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )