import calendar
import datetime
import os
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="S.I.C.P. - Control de Fichaje y Vacaciones DGFP",
    page_icon="⚡",
    layout="wide",
)

# ESTILO VISUAL MODO OSCURO + PANEL IZQUIERDO NEGRO + LUCES LED PARPADEANTES
st.markdown(
    """
    <style>
    .main {
        background-color: #0e1117;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        color: #ffffff;
    }
    
    /* Panel lateral de la izquierda NEGRO con título CONTROL HORARIO */
    [data-testid="stSidebar"] {
        background-color: #1a1c23 !important;
        border-right: 2px solid #333333;
    }
    [data-testid="stSidebar"] * {
        color: #e2e8f0 !important;
    }
    
    /* Barra superior institucional azul */
    .win-header {
        background: linear-gradient(90deg, #0b3c85 0%, #1d529e 100%);
        color: white;
        padding: 6px 12px;
        font-weight: bold;
        font-size: 13px;
        display: flex;
        justify-content: space-between;
        align-items: center;
        border-radius: 4px 4px 0 0;
        box-shadow: 0 2px 4px rgba(0,0,0,0.5);
    }
    
    .win-menu-bar {
        background-color: #262730;
        padding: 4px 10px;
        font-size: 12px;
        border-bottom: 1px solid #4f46e5;
        color: #d1d5db;
    }
    .win-menu-bar span {
        margin-right: 15px;
        cursor: pointer;
    }
    
    div.stButton > button {
        background-color: #262730 !important;
        color: #ffffff !important;
        border: 1px solid #4b5563 !important;
        border-radius: 3px !important;
        font-weight: 600;
        font-size: 13px;
    }
    div.stButton > button:hover {
        background-color: #374151 !important;
        border-color: #9ca3af !important;
    }

    .terminal-text {
        font-family: 'Courier New', Courier, monospace;
        color: #60a5fa;
        font-weight: bold;
        font-size: 13px;
        margin-bottom: 5px;
    }
    
    .badge-activo {
        background-color: #065f46;
        color: #d1fae5;
        padding: 4px 10px;
        border-radius: 12px;
        font-size: 12px;
        font-weight: bold;
        display: inline-block;
    }

    /* EFECTO LED PARPADEANTE (ROJO Y VERDE) */
    @keyframes blink-green {
        0% { opacity: 1; box-shadow: 0 0 10px #10b981; }
        50% { opacity: 0.4; box-shadow: 0 0 2px #10b981; }
        100% { opacity: 1; box-shadow: 0 0 10px #10b981; }
    }
    @keyframes blink-red {
        0% { opacity: 1; box-shadow: 0 0 10px #ef4444; }
        50% { opacity: 0.4; box-shadow: 0 0 2px #ef4444; }
        100% { opacity: 1; box-shadow: 0 0 10px #ef4444; }
    }

    .led-green {
        background-color: #065f46;
        color: #d1fae5;
        border: 2px solid #10b981;
        padding: 10px 20px;
        border-radius: 8px;
        font-size: 18px;
        font-weight: bold;
        text-align: center;
        animation: blink-green 1.5s infinite;
    }
    .led-red {
        background-color: #7f1d1d;
        color: #fee2e2;
        border: 2px solid #ef4444;
        padding: 10px 20px;
        border-radius: 8px;
        font-size: 18px;
        font-weight: bold;
        text-align: center;
        animation: blink-red 1.5s infinite;
    }

    /* TARJETAS MÉTRICAS ADAPTADAS A MODO OSCURO */
    .stMetric {
        background-color: #1f2937 !important;
        border: 1px solid #374151 !important;
        border-radius: 6px !important;
        padding: 15px !important;
        color: #ffffff !important;
    }
    .stMetric label {
        color: #9ca3af !important;
    }
    .stMetric [data-testid="stMetricValue"] {
        color: #ffffff !important;
    }
    </style>
""",
    unsafe_allow_html=True,
)

EXCEL_FILE = "Marcajes__2026_DGFP.xlsx"


def check_password():
  """Control de acceso con usuario aespiga y contraseña unirioja112 (permite Enter)."""
  if "password_correct" not in st.session_state:
    st.session_state["password_correct"] = False

  if not st.session_state["password_correct"]:
    st.markdown(
        '<div class="win-header">⚡ S.I.C.P. - Módulo de Fichaje y Registro de'
        ' Jornada [SESIÓN: RESTRINGIDA]</div>',
        unsafe_allow_html=True,
    )
    st.markdown("### 🔒 Acceso al Sistema")

    with st.form("login_form"):
      username_input = st.text_input("Usuario")
      password_input = st.text_input("Contraseña", type="password")
      submit_button = st.form_submit_button("Iniciar Sesión")

      if submit_button:
        if (
            username_input.strip().lower() == "aespiga"
            and password_input == "unirioja112"
        ):
          st.session_state["password_correct"] = True
          st.rerun()
        else:
          st.error(
              "❌ Usuario o contraseña incorrectos. (Usuario: aespiga / Clave:"
              " unirioja112)"
          )
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
  """Regla de verano: 15 Jun al 30 Sep = 6h 30m (6.5h).

  Resto del año (Horario de Invierno) = 7 horas (7.0h).
  """
  mes = fecha_dt.month
  dia = fecha_dt.day
  if (mes == 6 and dia >= 15) or (mes in [7, 8]) or (mes == 9 and dia <= 30):
    return 6.5, "6h 30m (Verano)"
  else:
    return 7.0, "7h 00m (Horario de Invierno / Estándar)"


# VENTANA EMERGENTE (POPUP) PARA REGISTRO / FICHAJE DIARIO
@st.dialog("✍️ Registrar / Modificar Fichaje Diario")
def popup_fichaje():
  st.write("Introduce los detalles de tu jornada laboral o incidencia:")

  fecha_elegida = st.date_input(
      "Selecciona la fecha:",
      value=datetime.date.today(),
      min_value=datetime.date(2026, 8, 1),
      max_value=datetime.date(2027, 12, 31),
  )

  horas_obj, desc_regimen = get_target_hours(fecha_elegida)
  st.info(f"🎯 **Jornada objetivo para este día**: {desc_regimen}")

  col_a, col_b = st.columns(2)
  with col_a:
    estado_dia = st.selectbox(
        "Tipo de día:",
        ["Trabajo", "Festivo", "Vacaciones", "Moscoso", "Permiso", "Baja Médica"],
    )
    h_inicio = st.time_input("Hora de Entrada", value=datetime.time(7, 30))
  with col_b:
    h_final = st.time_input(
        "Hora de Salida",
        value=datetime.time(14, 30) if horas_obj == 7.0 else datetime.time(14, 0),
    )
    observaciones = st.text_input("Observaciones")

  t1 = datetime.datetime.combine(datetime.date.today(), h_inicio)
  t2 = datetime.datetime.combine(datetime.date.today(), h_final)
  diff = t2 - t1
  h_trab = int(diff.total_seconds() // 3600)
  m_trab = int((diff.total_seconds() % 3600) // 60)

  diferencia_minutos = (h_trab * 60 + m_trab) - int(horas_obj * 60)

  st.write(
      f"⏱️ **Horas calculadas**: {h_trab}h {m_trab:02d}m (Objetivada en"
      f" {horas_obj}h)"
  )
  st.metric(
      "Balance diario",
      f"{h_trab}h {m_trab:02d}m",
      delta=f"{diferencia_minutos:+d} minutos respecto al objetivo",
  )

  if st.button("💾 Confirmar y Guardar Fichaje"):
    st.success(
        f"✅ ¡Fichaje registrado para el {fecha_elegida}! ({h_trab}h"
        f" {m_trab:02d}m)"
    )
    st.rerun()


# VENTANA EMERGENTE (POPUP) PARA VACACIONES Y MOSCOSOS CON SELECTOR
@st.dialog("🏖️ Solicitar / Registrar Vacaciones o Moscosos")
def popup_vacaciones():
  st.write(
      "Selecciona si se trata de un día de Vacaciones o un Moscoso e introduce"
      " las fechas:"
  )

  anno = st.radio("Año aplicable:", [2026, 2027], horizontal=True)

  tipo_permiso = st.selectbox(
      "Selecciona tipo de día:",
      ["🌴 Día de Vacaciones", "📑 Día Moscoso / Asuntos Propios"],
  )

  col1, col2 = st.columns(2)
  with col1:
    f_inicio = st.date_input(
        "Fecha de inicio",
        value=datetime.date(anno, 11, 16)
        if anno == 2026
        else datetime.date(2027, 7, 1),
    )
  with col2:
    f_fin = st.date_input(
        "Fecha de fin",
        value=datetime.date(anno, 11, 20)
        if anno == 2026
        else datetime.date(2027, 7, 5),
    )

  motivo = st.text_input("Observaciones o motivo (opcional)")

  dias_solicitados = (f_fin - f_inicio).days + 1

  st.write(
      f"📊 **Total días a computar**: **{dias_solicitados} día(s)** como"
      f" **{tipo_permiso}** para el año **{anno}**."
  )

  if st.button("💾 Guardar Permiso en Sistema"):
    st.success(
        f"✅ ¡Se han registrado {dias_solicitados} día(s) de {tipo_permiso} para"
        f" el año {anno} con éxito!"
    )
    st.rerun()


def main():
  if not check_password():
    return

  # BARRA SUPERIOR DE VENTANA INSTITUCIONAL
  st.markdown(
      """
        <div class="win-header">
            <span>⚡ S.I.C.P. - Módulo de Fichaje y Control Horario (v1.2.49) [SESIÓN: AESPIGA]</span>
            <span>_ &nbsp; ▫ &nbsp; X</span>
        </div>
        <div class="win-menu-bar">
            <span><b>Archivo</b></span>
            <span><b>Edición</b></span>
            <span><b>Ver</b></span>
            <span><b>Herramientas</b></span>
            <span><b>Seguridad</b></span>
            <span><b>Ayuda</b></span>
        </div>
    """,
      unsafe_allow_html=True,
  )

  # BOTONES DE LA TOOLBAR SUPERIOR
  col_tb1, col_tb2, col_tb3, col_tb4, _ = st.columns([1, 1.2, 1.2, 1.2, 4])
  with col_tb1:
    if st.button("📂 Nuevo Fichaje"):
      popup_fichaje()
  with col_tb2:
    if st.button("💾 Guardar"):
      st.toast("Guardado correctamente")
  with col_tb3:
    if st.button("🖨️ Imprimir"):
      st.toast("Enviando informe a la impresora...")
  with col_tb4:
    if st.button("🚪 Cerrar Sesión"):
      st.session_state["password_correct"] = False
      st.rerun()

  st.markdown("<br>", unsafe_allow_html=True)

  # ENCABEZADO PRINCIPAL CON SALDO LED PARPADEANTE A FECHA DE HOY
  col_tit1, col_tit2 = st.columns([2.5, 1.5])

  with col_tit1:
    st.title("CONTROL DE FICHAJE Y JORNADA")
    st.markdown(
        f"**Usuario:** aespiga | **Fecha Actual:**"
        f" `{datetime.date.today().strftime('%d/%m/%Y')}`"
    )

  with col_tit2:
    # Simulación de saldo a fecha de hoy (Ejemplo: -0:55 min o +1:30h)
    # Puedes ajustar este valor o conectarlo a tu cálculo exacto.
    saldo_minutos = -55  # Cambiar a positivo o negativo según saldo real

    if saldo_minutos < 0:
      horas_abs = abs(saldo_minutos) // 60
      mins_abs = abs(saldo_minutos) % 60
      texto_saldo = f"-{horas_abs}:{mins_abs:02d} min."
      st.markdown(
          f'<div style="text-align: center; margin-top: 10px;"><div'
          f' class="led-red">🔴 SALDO A HOY: {texto_saldo}</div></div>',
          unsafe_allow_html=True,
      )
    else:
      horas_abs = saldo_minutos // 60
      mins_abs = saldo_minutos % 60
      texto_saldo = f"+{horas_abs}:{mins_abs:02d}h."
      st.markdown(
          f'<div style="text-align: center; margin-top: 10px;"><div'
          f' class="led-green">🟢 SALDO A HOY: {texto_saldo}</div></div>',
          unsafe_allow_html=True,
      )

  st.markdown(
      '<div class="terminal-text">>> AVISO: Periodo de invierno activo (Jornada'
      " estándar de 7:00h). Del 15 Jun al 30 Sep rige jornada de 6:30h.</div>",
      unsafe_allow_html=True,
  )

  st.markdown("---")

  # MENÚ LATERAL - CONTROL HORARIO
  st.sidebar.markdown("## 🗂️ CONTROL HORARIO")
  menu = st.sidebar.radio(
      "Seleccionar módulo:",
      [
          "📅 Ver Calendario Mensual",
          "✍️ Fichaje y Registro Diario",
          "🏖️ Vacaciones y Moscosos",
      ],
  )

  base_sheets = load_base_sheets()

  if menu == "📅 Ver Calendario Mensual":
    st.header("📅 Calendario Mensual de Jornada")

    col_m1, col_m2 = st.columns(2)
    with col_m1:
      anno_sel = st.selectbox("Año a consultar", [2026, 2027], index=0)
    with col_m2:
      mes_sel = st.selectbox(
          "Mes",
          [
              "Enero",
              "Febrero",
              "Marzo",
              "Abril",
              "Mayo",
              "Junio",
              "Julio",
              "Agosto",
              "Septiembre",
              "Octubre",
              "Noviembre",
              "Diciembre",
          ],
          index=7,
      )

    meses_num = {
        "Enero": 1,
        "Febrero": 2,
        "Marzo": 3,
        "Abril": 4,
        "Mayo": 5,
        "Junio": 6,
        "Julio": 7,
        "Agosto": 8,
        "Septiembre": 9,
        "Octubre": 10,
        "Noviembre": 11,
        "Diciembre": 12,
    }
    m_num = meses_num[mes_sel]

    num_dias = calendar.monthrange(anno_sel, m_num)[1]
    fechas = [datetime.date(anno_sel, m_num, d) for d in range(1, num_dias + 1)]

    data_gen = []
    dias_semana_nombres = [
        "lunes",
        "martes",
        "miércoles",
        "jueves",
        "viernes",
        "sábado",
        "domingo",
    ]

    for f in fechas:
      d_sem = dias_semana_nombres[f.weekday()]
      fest_lab = "Festivo" if f.weekday() >= 5 else "Trabajo"
      obj_h, desc_obj = get_target_hours(f)

      data_gen.append({
          "Fecha": f.strftime("%Y-%m-%d"),
          "Día Semana": d_sem,
          "Laboral/Festivo": fest_lab,
          "Entrada": "07:30:00" if fest_lab == "Trabajo" else "",
          "Salida": (
              ("14:00:00" if obj_h == 6.5 else "14:30:00")
              if fest_lab == "Trabajo"
              else ""
          ),
          "Objetivo": desc_obj,
          "Observaciones": "",
      })

    st.dataframe(pd.DataFrame(data_gen), use_container_width=True)

  elif menu == "✍️ Fichaje y Registro Diario":
    st.header("✍️ Fichaje Diario y Control Horario")
    st.write(
        "Haz clic en el botón para abrir la ventana emergente de fichaje"
        " (configurada con el objetivo de invierno de **7 horas**):"
    )

    if st.button("➕ Abrir Ventana Emergente de Fichaje", type="primary"):
      popup_fichaje()

  elif menu == "🏖️ Vacaciones y Moscosos":
    st.header("🏖️ Gestor de Vacaciones y Moscosos")

    if st.button("➕ Abrir Ventana Emergente para Añadir Permiso / Moscoso"):
      popup_vacaciones()

    st.markdown("<br>", unsafe_allow_html=True)
    tab_2026, tab_2027 = st.tabs(
        ["📅 Resumen 2026 (Restantes: 3 Vacaciones)", "📅 Resumen 2027 (22/8)"]
    )

    with tab_2026:
      st.subheader("Saldos 2026")
      col_v1, col_v2, col_v3 = st.columns(3)
      with col_v1:
        st.metric(
            "Vacaciones 2026", "5 Disfrutadas / 8 Totales", "Restantes: 3 días"
        )
      with col_v2:
        st.metric(
            "Moscosos 2026", "3 Disfrutados / 3 Totales", "Restantes: 0 días"
        )
      with col_v3:
        st.metric("Cierre Año", "Noviembre / Diciembre", "Conforme")

      st.markdown("#### Registrados en 2026:")
      df_reg_2026 = pd.DataFrame({
          "Concepto": [
              "Vacaciones",
              "Vacaciones",
              "Vacaciones",
              "Vacaciones",
              "Vacaciones",
          ],
          "Fecha": [
              "16/11/2026",
              "17/11/2026",
              "18/11/2026",
              "19/11/2026",
              "20/11/2026",
          ],
          "Estado": [
              "Disfrutado",
              "Disfrutado",
              "Disfrutado",
              "Disfrutado",
              "Disfrutado",
          ],
      })
      st.dataframe(df_reg_2026, use_container_width=True)

    with tab_2027:
      st.subheader("Saldos 2027")
      col_71, col_72 = st.columns(2)
      with col_71:
        st.metric(
            "Vacaciones 2027", "Disponibles: 22 días", "22 días restantes"
        )
      with col_72:
        st.metric("Moscosos 2027", "Disponibles: 8 días", "8 días restantes")


if __name__ == "__main__":
  main()
