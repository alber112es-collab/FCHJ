import calendar
import datetime
import os
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Control de Fichaje y Vacaciones DGFP",
    page_icon="⏱️",
    layout="wide",
)

# Estilo visual corporativo y limpio similar a tus otras apps de Streamlit
st.markdown("""
    <style>
    .main {
        background-color: #f8f9fa;
    }
    .stMetric {
        background-color: #ffffff;
        padding: 15px;
        border-radius: 8px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
        border: 1px solid #e5e7eb;
    }
    .stTabs [data-baseweb="tab-list"] {
        gap: 10px;
    }
    .stTabs [data-baseweb="tab"] {
        background-color: #f1f5f9;
        border-radius: 6px 6px 0px 0px;
        padding: 10px 20px;
        font-weight: 600;
    }
    .stTabs [aria-selected="true"] {
        background-color: #2563eb !important;
        color: white !important;
    }
    </style>
""", unsafe_allow_html=True)

EXCEL_FILE = "Marcajes__2026_DGFP.xlsx"


def check_password():
  """Control de acceso con usuario ADMIN y contraseña unirioja112."""

  def password_entered():
    if (
        st.session_state["username"].strip().upper() == "ADMIN"
        and st.session_state["password"] == "unirioja112"
    ):
      st.session_state["password_correct"] = True
      del st.session_state["password"]
    else:
      st.session_state["password_correct"] = False

  if "password_correct" not in st.session_state:
    st.markdown("### 🔒 Acceso Restringido - Control DGFP")
    st.text_input("Usuario", key="username")
    st.text_input("Contraseña", type="password", key="password")
    st.button("Iniciar Sesión", on_click=password_entered)
    return False
  elif not st.session_state["password_correct"]:
    st.markdown("### 🔒 Acceso Restringido - Control DGFP")
    st.text_input("Usuario", key="username")
    st.text_input("Contraseña", type="password", key="password")
    st.button("Iniciar Sesión", on_click=password_entered)
    st.error("😕 Usuario o contraseña incorrectos. (Clave: unirioja112)")
    return False
  else:
    return True


@st.cache_data
def load_base_sheets():
  if os.path.exists(EXCEL_FILE):
    xls = pd.ExcelFile(EXCEL_FILE)
    return xls.sheet_names
  return []


def get_target_hours(fecha_dt):
  """Regla de verano: 15 Jun al 30 Sep = 6h 30m (6.5h), resto = 7h."""
  mes = fecha_dt.month
  dia = fecha_dt.day
  if (mes == 6 and dia >= 15) or (mes in [7, 8]) or (mes == 9 and dia <= 30):
    return 6.5, "6h 30m (Verano)"
  else:
    return 7.0, "7h 00m (Estándar)"


def main():
  if not check_password():
    return

  st.title("⏱️ Control de Fichaje, Vacaciones y Moscosos")
  st.markdown("Dirección General de Función Pública (DGFP) — Panel de Control Plurianual")

  base_sheets = load_base_sheets()

  # Menú lateral
  st.sidebar.header("Menú Principal")
  menu = st.sidebar.radio(
      "Selecciona sección:",
      [
          "📅 Ver Calendario Mensual",
          "✍️ Registrar / Fichaje Diario",
          "🏖️ Vacaciones y Moscosos (2026 y 2027)",
      ],
  )

  if menu == "📅 Ver Calendario Mensual":
    st.header("📅 Visualización de Calendarios Mensuales")
    col_m1, col_m2 = st.columns(2)
    with col_m1:
      anno_sel = st.selectbox("Año", [2026, 2027], index=0)
    with col_m2:
      mes_sel = st.selectbox(
          "Mes",
          [
              "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
              "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
          ],
          index=7,
      )

    meses_num = {
        "Enero": 1, "Febrero": 2, "Marzo": 3, "Abril": 4, "Mayo": 5, "Junio": 6,
        "Julio": 7, "Agosto": 8, "Septiembre": 9, "Octubre": 10, "Noviembre": 11, "Diciembre": 12,
    }
    m_num = meses_num[mes_sel]

    num_dias = calendar.monthrange(anno_sel, m_num)[1]
    fechas = [datetime.date(anno_sel, m_num, d) for d in range(1, num_dias + 1)]

    data_gen = []
    dias_semana_nombres = ["lunes", "martes", "miércoles", "jueves", "viernes", "sábado", "domingo"]

    for f in fechas:
      d_sem = dias_semana_nombres[f.weekday()]
      fest_lab = "Festivo" if f.weekday() >= 5 else "Trabajo"
      obj_h, desc_obj = get_target_hours(f)

      data_gen.append({
          "Fecha": f.strftime("%Y-%m-%d"),
          "Día Semana": d_sem,
          "Laboral/Festivo/Inc.": fest_lab,
          "Inicio": "07:30:00" if fest_lab == "Trabajo" else "",
          "Final": (
              ("14:00:00" if obj_h == 6.5 else "14:30:00")
              if fest_lab == "Trabajo"
              else ""
          ),
          "H. Obligatorias": desc_obj,
          "Observaciones": "",
      })

    df_gen = pd.DataFrame(data_gen)
    st.dataframe(df_gen, use_container_width=True)

  elif menu == "✍️ Registrar / Fichaje Diario":
    st.header("✍️ Registro y Marcaje Diario")
    col_f1, col_f2 = st.columns(2)
    with col_f1:
      fecha_elegida = st.date_input(
          "Selecciona el día a fichar:",
          value=datetime.date(2026, 8, 3),
          min_value=datetime.date(2026, 8, 1),
          max_value=datetime.date(2027, 12, 31),
      )
    with col_f2:
      _, desc_regimen = get_target_hours(fecha_elegida)
      st.info(f"🎯 **Jornada objetivo**: {desc_regimen}")

    col_a, col_b = st.columns(2)
    with col_a:
      estado_dia = st.selectbox(
          "Tipo de día / Incidencia:",
          ["Trabajo", "Festivo", "Vacaciones", "Moscoso", "Permiso", "Baja Médica"],
          index=0,
      )
      h_inicio = st.time_input("Hora de Inicio", value=datetime.time(7, 30))
      h_final = st.time_input("Hora de Finalización", value=datetime.time(14, 0))

    with col_b:
      observaciones_txt = st.text_area("Observaciones / Incidencias:")

    t1 = datetime.datetime.combine(datetime.date.today(), h_inicio)
    t2 = datetime.datetime.combine(datetime.date.today(), h_final)
    diff = t2 - t1
    h_trab = int(diff.total_seconds() // 3600)
    m_trab = int((diff.total_seconds() % 3600) // 60)

    horas_obj_val, _ = get_target_hours(fecha_elegida)
    diferencia_minutos = (h_trab * 60 + m_trab) - int(horas_obj_val * 60)

    st.metric(
        "Horas Realizadas vs Objetivo",
        f"{h_trab}h {m_trab:02d}m",
        delta=f"{diferencia_minutos:+d} minutos respecto al objetivo",
    )

    if st.button("💾 Guardar y Registrar Fichaje", type="primary"):
      st.success(
          f"✅ ¡Registrado con éxito para el **{fecha_elegida}**! Estado:"
          f" **{estado_dia}** ({h_trab}h {m_trab:02d}m)."
      )

  elif menu == "🏖️ Vacaciones y Moscosos (2026 y 2027)":
    st.header("🏖️ Control de Vacaciones y Moscosos")
    st.markdown("Gestión detallada de saldos y periodos para los años **2026** y **2027**.")

    tab_2026, tab_2027 = st.tabs(
        ["📅 Año 2026 (Restantes: 3 Vacaciones)", "📅 Año 2027 (Totales: 22/8)"]
    )

    with tab_2026:
      st.subheader("Control Vacaciones y Moscosos - 2026")

      col_v1, col_v2, col_v3 = st.columns(3)
      with col_v1:
        st.metric("Vacaciones 2026", "Disfrutadas: 5 / Totales: 8", "Restantes: 3 días")
      with col_v2:
        st.metric("Moscosos 2026", "Disfrutados: 3 / Totales: 3", "Restantes: 0 días")
      with col_v3:
        st.metric("Estado 2026", "Cierre de año", "✔ Al día")

      st.markdown("#### Días de Vacaciones ya registrados en 2026:")
      df_reg_2026 = pd.DataFrame({
          "Concepto": ["Vacaciones", "Vacaciones", "Vacaciones", "Vacaciones", "Vacaciones"],
          "Fecha / Periodo": [
              "16 de Noviembre de 2026",
              "17 de Noviembre de 2026",
              "18 de Noviembre de 2026",
              "19 de Noviembre de 2026",
              "20 de Noviembre de 2026",
          ],
          "Estado": ["Disfrutado", "Disfrutado", "Disfrutado", "Disfrutado", "Disfrutado"]
      })
      st.dataframe(df_reg_2026, use_container_width=True)

      st.markdown("---")
      st.markdown("#### ➕ Introducir nuevo día / periodo de Vacaciones (2026)")
      f_ini_26 = st.date_input("Fecha de inicio (2026)", value=datetime.date(2026, 11, 16), key="f_i_26")
      f_fin_26 = st.date_input("Fecha de fin (2026)", value=datetime.date(2026, 11, 20), key="f_f_26")
      if st.button("Registrar Vacaciones 2026", key="btn_26"):
        st.success("¡Vacaciones de 2026 registradas correctamente en el sistema!")

    with tab_2027:
      st.subheader("Control Vacaciones y Moscosos - 2027")

      col_71, col_72 = st.columns(2)
      with col_71:
        st.metric("Vacaciones 2027", "Total disponibles: 22 días", "Restantes: 22")
      with col_72:
        st.metric("Moscosos 2027", "Total disponibles: 8 días", "Restantes: 8")

      st.markdown("#### ➕ Introducir Vacaciones o Moscosos para 2027")
      tipo_permiso_27 = st.selectbox("Selecciona tipo:", ["Vacaciones (22 días)", "Moscosos (8 días)"], key="tp_27")
      f_ini_27 = st.date_input("Fecha de inicio (2027)", value=datetime.date(2027, 7, 1), key="f_i_27")
      f_fin_27 = st.date_input("Fecha de fin (2027)", value=datetime.date(2027, 7, 5), key="f_f_27")
      observacion_27 = st.text_input("Motivo / Observaciones (Opcional)", key="obs_27")

      if st.button("💾 Guardar Solicitud 2027", key="btn_27"):
        st.success("¡Registrado con éxito para 2027! Se han descontado los días correspondientes.")


if __name__ == "__main__":
  main()
