from datetime import datetime
import os
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Control de Retardos y Faltas", page_icon="📱", layout="centered"
)

EXCEL_FILE = "LISTAS OCTUBRE (1).xlsx"
REGISTROS_FILE = "registros_asistencia.csv"


# --- CARGAR LISTA DE ESTUDIANTES DESDE EL EXCEL ---
@st.cache_data
def cargar_estudiantes():
  try:
    # Usamos la hoja maestra 'proc9324961' que contiene todos los estudiantes y grupos
    df_proc = pd.read_excel(EXCEL_FILE, sheet_name="proc9324961")

    # Construir el nombre completo limpiando espacios nulos
    df_proc["NOMBRE_ESTUDIANTES"] = (
        df_proc["NOMBRE1"].fillna("")
        + " "
        + df_proc["NOMBRE2"].fillna("")
        + " "
        + df_proc["APELLIDO1"].fillna("")
        + " "
        + df_proc["APELLIDO2"].fillna("")
    )
    df_proc["NOMBRE_ESTUDIANTES"] = (
        df_proc["NOMBRE_ESTUDIANTES"]
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
    )
    df_proc["CURSO"] = "Curso " + df_proc["GRUPO"].astype(str)

    # Seleccionar columnas requeridas
    estudiantes = df_proc[["CURSO", "NOMBRE_ESTUDIANTES"]].dropna(
        subset=["NOMBRE_ESTUDIANTES"]
    )
    return estudiantes
  except Exception as e:
    st.error(f"Error al leer el archivo Excel: {e}")
    return pd.DataFrame(columns=["CURSO", "NOMBRE_ESTUDIANTES"])


df_estudiantes = cargar_estudiantes()

# --- CARGAR O INICIALIZAR REGISTROS HISTÓRICOS ---
if os.path.exists(REGISTROS_FILE):
  df_registros = pd.read_csv(REGISTROS_FILE)
else:
  df_registros = pd.DataFrame(
      columns=["FECHA", "TIPO", "CURSO", "NOMBRE_ESTUDIANTE", "MES", "ANIO"]
  )

st.title("📱 Control de Asistencia Escolar")
st.write("Lleva el registro de retardos y fallas acumulado mes a mes.")

# --- MENÚ DE NAVEGACIÓN MÓVIL ---
menu = st.radio(
    "Selecciona una opción:",
    ["📝 Registrar Incidencia", "📊 Ver Acumulados y Reportes"],
    horizontal=True,
)

# --- 1. REGISTRAR DESDE EL CELULAR ---
if menu == "📝 Registrar Incidencia":
  st.subheader("Registrar Retardo o Falla")

  if df_estudiantes.empty:
    st.warning("No se encontraron estudiantes en el archivo Excel.")
  else:
    # Seleccionar Curso
    cursos_disponibles = sorted(df_estudiantes["CURSO"].unique().tolist())
    curso_sel = st.selectbox("1. Selecciona el CURSO:", cursos_disponibles)

    # Filtrar estudiantes de ese curso
    estudiantes_filtrados = sorted(
        df_estudiantes[df_estudiantes["CURSO"] == curso_sel][
            "NOMBRE_ESTUDIANTES"
        ].tolist()
    )
    estudiante_sel = st.selectbox(
        "2. Selecciona el ESTUDIANTE:", estudiantes_filtrados
    )

    # Tipo de incidencia
    tipo_sel = st.radio(
        "3. Tipo de incidencia:", ["Retardo", "Falla"], horizontal=True
    )

    fecha_hoy = datetime.now().strftime("%Y-%m-%d")
    mes_actual = datetime.now().strftime("%B")  # Nombre del mes
    anio_actual = datetime.now().strftime("%Y")

    st.info(f"📅 Fecha de registro: {fecha_hoy} ({mes_actual} {anio_actual})")

    if st.button("Guardar Registro 💾", type="primary", use_container_width=True):
      nuevo_reg = pd.DataFrame({
          "FECHA": [fecha_hoy],
          "TIPO": [tipo_sel],
          "CURSO": [curso_sel],
          "NOMBRE_ESTUDIANTE": [estudiante_sel],
          "MES": [mes_actual],
          "ANIO": [anio_actual],
      })

      # Agregar al DataFrame existente y guardar en CSV persistente
      df_registros = pd.concat([df_registros, nuevo_reg], ignore_index=True)
      df_registros.to_csv(REGISTROS_FILE, index=False)

      st.success(
          f"✅ ¡Guardado! Se registró un **{tipo_sel}** para"
          f" **{estudiante_sel}** ({curso_sel})."
      )

# --- 2. VER ACUMULADOS MES A MES Y AÑO ---
elif menu == "📊 Ver Acumulados y Reportes":
  st.subheader("📊 Reportes y Acumulados")

  if df_registros.empty:
    st.info("Aún no hay registros guardados en el sistema.")
  else:
    # Filtro por año
    anios_disponibles = sorted(df_registros["ANIO"].unique().tolist())
    anio_filtro = st.selectbox("Filtrar por Año:", anios_disponibles)

    df_anio = df_registros[df_registros["ANIO"] == str(anio_filtro)]

    # Filtro por mes (opcional o vista general)
    meses_disponibles = ["TODOS LOS MESES"] + sorted(
        df_anio["MES"].unique().tolist()
    )
    mes_filtro = st.selectbox("Filtrar por Mes:", meses_disponibles)

    if mes_filtro != "TODOS LOS MESES":
      df_filtrado = df_anio[df_anio["MES"] == mes_filtro]
      st.markdown(f"### Acumulado para el mes de: {mes_filtro} ({anio_filtro})")
    else:
      df_filtrado = df_anio
      st.markdown(f"### Acumulado Anual Total ({anio_filtro})")

    if df_filtrado.empty:
      st.warning("No hay registros para los filtros seleccionados.")
    else:
      # Tabla resumen agrupada por Curso y Estudiante
      resumen = (
          df_filtrado.pivot_table(
              index=["CURSO", "NOMBRE_ESTUDIANTE"],
              columns="TIPO",
              values="FECHA",
              aggfunc="count",
              fill_value=0,
          )
          .reset_index()
      )

      # Asegurar columnas de Retardo y Falla si no existen
      for col in ["Retardo", "Falla"]:
        if col not in resumen.columns:
          resumen[col] = 0

      # Calcular total de incidencias
      resumen["TOTAL_INCIDENCIAS"] = resumen["Retardo"] + resumen["Falla"]
      resumen = resumen.sort_values(
          by="TOTAL_INCIDENCIAS", ascending=False
      ).reset_index(drop=True)

      st.dataframe(resumen, use_container_width=True)

      # Botón para descargar reporte en Excel
      output = io.BytesIO() if "io" in globals() else None
      import io

      output = io.BytesIO()
      with pd.ExcelWriter(output, engine="openpyxl") as writer:
        resumen.to_excel(
            writer, index=False, sheet_name="Acumulado_Asistencia"
        )
        df_filtrado.to_excel(
            writer, index=False, sheet_name="Detalle_Completo"
        )
      excel_data = output.getvalue()

      st.download_button(
          label="📥 Descargar Reporte en Excel",
          data=excel_data,
          file_name=f"reporte_asistencia_{anio_filtro}.xlsx",
          mime=(
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          ),
          use_container_width=True,
      )

      with st.expander("Ver Historial Detallado"):
        st.dataframe(df_filtrado, use_container_width=True) 
