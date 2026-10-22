import datetime
import os
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Control de Fichaje DGFP", page_icon="⏱️", layout="wide"
)

EXCEL_FILE = "Marcajes__2026_DGFP.xlsx"


@st.cache_data
def load_data():
  if os.path.exists(EXCEL_FILE):
    xls = pd.ExcelFile(EXCEL_FILE)
    sheets = xls.sheet_names
    return sheets
  return []


def main():
  st.title("⏱️ Aplicación de Control de Fichaje y Calendario")
  st.markdown(
      "Gestiona tu jornada laboral diaria (**7 horas diarias** / u objetivo"
      " mensual) y consulta tu calendario."
  )

  sheets = load_data()

  if not sheets:
    st.error(
        f"No se encuentra el archivo '{EXCEL_FILE}' en el directorio."
        " Asegúrate de subirlo a GitHub junto con este script."
    )
    return

  # Sidebar para navegación
  st.sidebar.header("Menú de Navegación")
  menu = st.sidebar.radio(
      "Selecciona una opción:",
      [
          "📅 Vista de Calendario y Meses",
          "✍️ Registro / Fichaje Diario",
          "🏖️ Vacaciones y Moscosos",
      ],
  )

  # Seleccionar hoja de mes activa (excluyendo la de moscosos/vacaciones para el calendario diario si procede)
  meses_sheets = [s for s in sheets if "MOSCOSOS" not in s]

  if menu == "📅 Vista de Calendario y Meses":
    st.header("📅 Visualización del Calendario Mensual")

    selected_sheet = st.selectbox("Selecciona el mes a consultar:", meses_sheets)

    if selected_sheet:
      df_mes = pd.read_excel(EXCEL_FILE, sheet_name=selected_sheet, skiprows=3)
      # Limpiar columnas vacías
      df_mes = df_mes.loc[:, ~df_mes.columns.str.contains("^Unnamed")]

      st.subheader(f"Resumen del mes: {selected_sheet}")

      # Mostrar métricas rápidas si existen
      col1, col2, col3 = st.columns(3)
      with col1:
        st.metric("Jornada diaria objetivo", "7h 00m")
      with col2:
        if "Laboral/Festivo/Inc." in df_mes.columns:
          dias_trabajo = (
              df_mes["Laboral/Festivo/Inc."] == "Trabajo"
          ).sum()
          st.metric("Días Laborables", int(dias_trabajo))
      with col3:
        if "HH.TT." in df_mes.columns:
          st.info("Consulta el detalle diario en la tabla inferior.")

      # Mostrar tabla interactiva
      st.dataframe(df_mes, use_container_width=True)

  elif menu == "✍️ Registro / Fichaje Diario":
    st.header("✍️ Registro y Modificación del Fichaje Diario")
    st.markdown(
        "Introduce o modifica las horas de entrada y salida para registrar tu"
        " jornada (Objetivo: **7 horas al día**)."
    )

    selected_sheet = st.selectbox(
        "Selecciona el mes para registrar:", meses_sheets, key="fichaje_mes"
    )

    if selected_sheet:
      df_mes = pd.read_excel(EXCEL_FILE, sheet_name=selected_sheet, skiprows=3)
      # Limpiar columnas vacías en los nombres de columnas lógicos
      # Buscamos las columnas clave: Fecha, Día Semana, Laboral/Festivo/Inc., Inicio, Final, Observaciones

      # Filtramos filas que tengan fecha válida
      df_mes["Fecha_dt"] = pd.to_datetime(df_mes["Fecha"], errors="coerce")
      df_valid = df_mes.dropna(subset=["Fecha_dt"]).copy()

      if not df_valid.empty:
        # Selector de fecha dentro del mes
        fechas_disponibles = df_valid["Fecha_dt"].dt.date.tolist()
        fecha_elegida = st.selectbox(
            "Selecciona el día a fichar:", fechas_disponibles
        )

        # Fila correspondiente
        fila_idx = df_valid[
            df_valid["Fecha_dt"].dt.date == fecha_elegida
        ].index[0]

        row_data = df_mes.loc[fila_idx]

        col_a, col_b = st.columns(2)
        with col_a:
          estado_actual = st.selectbox(
              "Tipo de día / Estado:",
              ["Trabajo", "Festivo", "Vacaciones", "Moscosos", "Permiso", "Baja"],
              index=0,
          )
          hora_inicio = st.time_input(
              "Hora de Inicio",
              value=datetime.time(7, 30),
          )
          hora_final = st.time_input(
              "Hora de Finalización",
              value=datetime.time(14, 30),
          )

        with col_b:
          observacion_txt = st.text_area(
              "Observaciones / Incidencias:",
              value=str(row_data.get("Observaciones", ""))
              if pd.notna(row_data.get("Observaciones", ""))
              else "",
          )

        # Cálculo automático de horas
        # 7 horas obligatorias por defecto
        t1 = datetime.datetime.combine(datetime.date.today(), hora_inicio)
        t2 = datetime.datetime.combine(datetime.date.today(), hora_final)
        diff = t2 - t1
        horas_trabajadas_seg = diff.total_seconds()
        h_trab = int(horas_trabajadas_seg // 3600)
        m_trab = int((horas_trabajadas_seg % 3600) // 60)

        st.metric(
            "Horas calculadas para este día", f"{h_trab}h {m_trab:02d}m"
        )

        if st.button("💾 Guardar Fichaje (Simulación en Sesión)"):
          st.success(
              f"¡Fichaje guardado correctamente para el día {fecha_elegida}! "
              f"({h_trab}h {m_trab:02d}m registradas)."
          )
          st.info(
              "Nota: Para persistir los cambios permanentemente en el archivo"
              " Excel de tu repositorio, puedes conectar la app a una base de"
              " datos (como SQLite/Supabase) o actualizar directamente el"
              " Excel en GitHub."
          )
      else:
        st.warning(
            "No se han encontrado fechas válidas estructuradas en esta hoja."
        )

  elif menu == "🏖️ Vacaciones y Moscosos":
    st.header("🏖️ Control de Vacaciones, Moscosos y Permisos")
    if "MOSCOSOS-VACACIONES - 2026" in sheets:
      df_vac = pd.read_excel(
          EXCEL_FILE, sheet_name="MOSCOSOS-VACACIONES - 2026"
      )
      st.dataframe(df_vac, use_container_width=True)
    else:
      st.info("Pestaña de vacaciones no detectada exactamente con ese nombre.")


if __name__ == "__main__":
  main()