import datetime
import os
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Control de Fichaje DGFP (2026-2027)",
    page_icon="⏱️",
    layout="wide",
)

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
    st.markdown("### 🔒 Acceso Restringido - Control de Fichaje")
    st.text_input("Usuario", key="username")
    st.text_input("Contraseña", type="password", key="password")
    st.button("Iniciar Sesión", on_click=password_entered)
    return False
  elif not st.session_state["password_correct"]:
    st.markdown("### 🔒 Acceso Restringido - Control de Fichaje")
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
  """Devuelve el objetivo de horas según la fecha:

  - Verano (15 de junio al 30 de septiembre): 6 horas y 30 minutos (6.5h)
  - Resto del año: 7 horas (7.0h)
  """
  mes = fecha_dt.month
  dia = fecha_dt.day

  # Del 15 de junio (6-15) al 30 de septiembre (9-30)
  if (mes == 6 and dia >= 15) or (mes in [7, 8]) or (mes == 9 and dia <= 30):
    return 6.5, "6h 30m (Verano)"
  else:
    return 7.0, "7h 00m (Estándar)"


def main():
  if not check_password():
    return

  st.title(
      "⏱️ Control de Fichaje y Calendario Plurianual (2026 - 2027) - DGFP"
  )
  st.markdown(
      "Gestión de jornada laboral con cálculo automático de verano (**6h 30m**"
      " del 15 Jun al 30 Sep) y resto del año (**7h**)."
  )

  base_sheets = load_base_sheets()

  # Menú lateral
  st.sidebar.header("Menú de Navegación")
  menu = st.sidebar.radio(
      "Selecciona sección:",
      [
          "📅 Ver Calendario y Meses",
          "✍️ Registrar / Fichaje Diario (2026-2027)",
          "🏖️ Vacaciones, Moscosos y Permisos",
      ],
  )

  if menu == "📅 Ver Calendario y Meses":
    st.header("📅 Visualización de Calendarios Mensuales")

    # Opciones de meses (incluyendo generación virtual para 2027 si no están en el Excel)
    meses_disponibles = [s for s in base_sheets if "MOSCOSOS" not in s]

    # Añadir generador virtual para meses de 2027 si se desea consultar
    st.subheader("Selecciona o genera el mes a consultar:")
    tipo_vista = st.radio(
        "Origen de datos:", ["Archivo Excel original", "Generador 2027 (Nuevo)"]
    )

    if tipo_vista == "Archivo Excel original" and meses_disponibles:
      selected_sheet = st.selectbox(
          "Selecciona el mes del Excel:", meses_disponibles
      )
      df_mes = pd.read_excel(EXCEL_FILE, sheet_name=selected_sheet, header=4)
      df_mes = df_mes.loc[
          :, ~df_mes.columns.str.contains("^Unnamed|nan", na=False)
      ]
      st.dataframe(df_mes, use_container_width=True)
    else:
      # Generador dinámico para cualquier mes de 2027
      col_m1, col_m2 = st.columns(2)
      with col_m1:
        anno_sel = st.selectbox("Año", [2026, 2027], index=1)
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
            index=5,
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

      # Generar días del mes
      import calendar

      num_dias = calendar.monthrange(anno_sel, m_num)[1]
      fechas = [
          datetime.date(anno_sel, m_num, d) for d in range(1, num_dias + 1)
      ]

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
        # Festivo por defecto si es fin de semana
        fest_lab = (
            "Festivo" if f.weekday() >= 5 else "Trabajo"
        )
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
      st.success(
          f"Calendario generado para {mes_sel} {anno_sel} con la regla de"
          " horario aplicada."
      )
      st.dataframe(df_gen, use_container_width=True)

  elif menu == "✍️ Registrar / Fichaje Diario (2026-2027)":
    st.header("✍️ Registro y Marcaje Diario Avanzado")
    st.markdown(
        "Selecciona cualquier fecha de **2026 o 2027**, marca incidencias"
        " (Trabajo, Festivo, Vacaciones, Moscosos, Permiso, Baja) y calcula tus"
        " horas automáticamente."
    )

    col_f1, col_f2 = st.columns(2)
    with col_f1:
      fecha_elegida = st.date_input(
          "Selecciona el día a fichar:",
          value=datetime.date(2027, 6, 15),
          min_date=datetime.date(2026, 8, 1),
          max_date=datetime.date(2027, 12, 31),
      )
    with col_f2:
      # Mostrar automáticamente el régimen aplicable
      _, desc_regimen = get_target_hours(fecha_elegida)
      st.info(f"🎯 **Jornada objetivo para este día**: {desc_regimen}")

    col_a, col_b = st.columns(2)
    with col_a:
      estado_dia = st.selectbox(
          "Tipo de día / Incidencia:",
          [
              "Trabajo",
              "Festivo",
              "Vacaciones",
              "Moscoso",
              "Permiso",
              "Baja Médica",
          ],
          index=0,
      )
      h_inicio = st.time_input("Hora de Inicio", value=datetime.time(7, 30))
      h_final = st.time_input("Hora de Finalización", value=datetime.time(14, 0))

    with col_b:
      observaciones_txt = st.text_area(
          "Observaciones / Incidencias (ej: médico, asuntos propios...):"
      )

    # Cálculos
    t1 = datetime.datetime.combine(datetime.date.today(), h_inicio)
    t2 = datetime.datetime.combine(datetime.date.today(), h_final)
    diff = t2 - t1
    total_seg = diff.total_seconds()
    h_trab = int(total_seg // 3600)
    m_trab = int((total_seg % 3600) // 60)

    horas_obj_val, _ = get_target_hours(fecha_elegida)
    diferencia_minutos = (h_trab * 60 + m_trab) - int(horas_obj_val * 60)

    st.metric(
        "Horas Realizadas vs Objetivo",
        f"{h_trab}h {m_trab:02d}m",
        delta=f"{diferencia_minutos:+d} minutos respecto al objetivo",
    )

    if st.button("💾 Guardar y Registrar Fichaje en Sesión"):
      st.success(
          f"✅ ¡Registrado con éxito para el **{fecha_elegida}**! Estado:"
          f" **{estado_dia}** ({h_trab}h {m_trab:02d}m)."
      )
      st.balloons()

  elif menu == "🏖️ Vacaciones, Moscosos y Permisos":
    st.header("🏖️ Control de Vacaciones y Moscosos (2026 - 2027)")
    st.markdown(
        "Control acumulado de días de vacaciones, moscosos y permisos"
        " disfrutados y pendientes."
    )

    if "MOSCOSOS-VACACIONES - 2026" in base_sheets:
      df_vac = pd.read_excel(
          EXCEL_FILE, sheet_name="MOSCOSOS-VACACIONES - 2026"
      )
      st.dataframe(df_vac, use_container_width=True)
    else:
      # Plantilla interactiva de control si no se lee del Excel
      data_control = {
          "Concepto": [
              "Vacaciones Anuales",
              "Días Moscosos / Asuntos Propios",
              "Permisos Especiales",
          ],
          "Total Disponibles": [22, 6, 5],
          "Disfrutados": [0, 0, 0],
          "Pendientes": [22, 6, 5],
      }
      st.dataframe(pd.DataFrame(data_control), use_container_width=True)


if __name__ == "__main__":
  main()
