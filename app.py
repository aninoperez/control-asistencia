from datetime import datetime
import os
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Control de Asistencia - I.E.D. Guillermo León Valencia",
    page_icon="📱",
    layout="centered",
)

EXCEL_FILE = "LISTAS OCTUBRE (1).xlsx"
REGISTROS_FILE = "registros_asistencia.csv"

# Diccionario de meses en español
MESES_ES = {
    1: "Enero",
    2: "Febrero",
    3: "Marzo",
    4: "Abril",
    5: "Mayo",
    6: "Junio",
    7: "Julio",
    8: "Agosto",
    9: "Septiembre",
    10: "Octubre",
    11: "Noviembre",
    12: "Diciembre",
}


# --- CARGAR LISTA DE ESTUDIANTES DIRECTAMENTE DE LAS PESTAÑAS DEL EXCEL ---
@st.cache_data
def cargar_estudiantes():
  try:
    xls = pd.ExcelFile(EXCEL_FILE)
    hojas_grados = [s for s in xls.sheet_names if s != "proc9324961"]
    master_list = []

    for sheet in hojas_grados:
      df = pd.read_excel(EXCEL_FILE, sheet_name=sheet)
      header_idx = None
      for idx, row in df.iterrows():
        row_str = str(row.values).upper()
        if "APELLIDO" in row_str or "NOMBRE" in row_str:
          header_idx = idx
          break

      if header_idx is not None:
        df.columns = df.iloc[header_idx]
        df = df.iloc[header_idx + 1 :].copy()
        df = df.loc[:, df.columns.notna()]

        n1 = next((c for c in df.columns if "NOMBRE1" in str(c).upper()), None)
        n2 = next((c for c in df.columns if "NOMBRE2" in str(c).upper()), None)
        a1 = next((c for c in df.columns if "APELLIDO1" in str(c).upper()), None)
        a2 = next((c for c in df.columns if "APELLIDO2" in str(c).upper()), None)

        if n1 and a1:
          df["NOMBRE_ESTUDIANTE"] = (
              df[n1].fillna("").astype(str)
              + " "
              + (df[n2].fillna("").astype(str) if n2 else "")
              + " "
              + df[a1].fillna("").astype(str)
              + " "
              + (df[a2].fillna("").astype(str) if a2 else "")
          )
          df["NOMBRE_ESTUDIANTE"] = (
              df["NOMBRE_ESTUDIANTE"]
              .str.replace(r"\s+", " ", regex=True)
              .str.strip()
          )
          df["CURSO"] = sheet

          sub_df = df[["CURSO", "NOMBRE_ESTUDIANTE"]].dropna(
              subset=["NOMBRE_ESTUDIANTE"]
          )
          master_list.append(sub_df)

    if master_list:
      return pd.concat(master_list, ignore_index=True)
    else:
      return pd.DataFrame(columns=["CURSO", "NOMBRE_ESTUDIANTE"])
  except Exception as e:
    st.error(f"Error al leer el archivo Excel: {e}")
    return pd.DataFrame(columns=["CURSO", "NOMBRE_ESTUDIANTE"])


df_estudiantes = cargar_estudiantes()

# --- CARGAR O INICIALIZAR REGISTROS HISTÓRICOS ---
if os.path.exists(REGISTROS_FILE):
  df_registros = pd.read_csv(REGISTROS_FILE)
  if "HORA" not in df_registros.columns:
    df_registros["HORA"] = "00:00"
  if "MES" not in df_registros.columns or "ANIO" not in df_registros.columns:
    df_registros = pd.DataFrame(
        columns=[
            "FECHA",
            "HORA",
            "TIPO",
            "CURSO",
            "NOMBRE_ESTUDIANTE",
            "MES",
            "ANIO",
        ]
    )
else:
  df_registros = pd.DataFrame(
      columns=[
          "FECHA",
          "HORA",
          "TIPO",
          "CURSO",
          "NOMBRE_ESTUDIANTE",
          "MES",
          "ANIO",
      ]
  )

# --- TÍTULO OFICIAL ---
st.title("CONTROL DE ASISTENCIA ESCOLAR")
st.subheader("COLEGIO GUILLERMO LEÓN VALENCIA I.E.D. - ENGATIVÁ")
st.write(
    "Lleva el control de retardos y fallas acumulado mes a mes desde tu"
    " celular."
)

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
    orden_cursos = [
        "JARDÍN",
        "TRANSICIÓN",
        "PRIMERO",
        "SEGUNDO",
        "TERCERO",
        "CUARTO",
        "QUINTO",
        "SEXTO",
        "701",
        "702",
        "OCTAVO",
        "901",
        "902",
        "DÉCIMO",
        "ONCE",
    ]
    cursos_disponibles = [
        c
        for c in orden_cursos
        if c in df_estudiantes["CURSO"].unique().tolist()
    ]
    otros = [
        c
        for c in df_estudiantes["CURSO"].unique().tolist()
        if c not in orden_cursos
    ]
    cursos_disponibles.extend(otros)

    curso_sel = st.selectbox("1. Selecciona el CURSO:", cursos_disponibles)

    estudiantes_filtrados = sorted(
        df_estudiantes[df_estudiantes["CURSO"] == curso_sel][
            "NOMBRE_ESTUDIANTE"
        ].tolist()
    )
    estudiante_sel = st.selectbox(
        "2. Selecciona el ESTUDIANTE:", estudiantes_filtrados
    )

    tipo_sel = st.radio(
        "3. Tipo de incidencia:", ["Retardo", "Falla"], horizontal=True
    )

    now = datetime.now()
    fecha_hoy = now.strftime("%Y-%m-%d")
    hora_actual = now.strftime("%H:%M")
    mes_actual = MESES_ES.get(now.month, "Desconocido")
    anio_actual = str(now.year)

    st.info(
        f"📅 Fecha y Hora de registro: {fecha_hoy} a las {hora_actual} ({mes_actual}"
        f" {anio_actual})"
    )

    if st.button("Guardar Registro 💾", type="primary", use_container_width=True):
      nuevo_reg = pd.DataFrame({
          "FECHA": [fecha_hoy],
          "HORA": [hora_actual],
          "TIPO": [tipo_sel],
          "CURSO": [curso_sel],
          "NOMBRE_ESTUDIANTE": [estudiante_sel],
          "MES": [mes_actual],
          "ANIO": [anio_actual],
      })

      df_registros = pd.concat([df_registros, nuevo_reg], ignore_index=True)
      df_registros.to_csv(REGISTROS_FILE, index=False)

      st.success(
          f"✅ ¡Guardado! Se registró un **{tipo_sel}** para"
          f" **{estudiante_sel}** ({curso_sel}) a las {hora_actual}."
      )

# --- 2. VER ACUMULADOS MES A MES Y AÑO ---
elif menu == "📊 Ver Acumulados y Reportes":
  st.subheader("📊 Reportes y Acumulados")

  if df_registros.empty:
    st.info("Aún no hay registros guardados en el sistema.")
  else:
    df_registros["ANIO"] = df_registros["ANIO"].astype(str)
    df_registros["MES"] = df_registros["MES"].astype(str)

    anios_disponibles = sorted(df_registros["ANIO"].unique().tolist())
    anio_filtro = st.selectbox("Filtrar por Año:", anios_disponibles)

    df_anio = df_registros[df_registros["ANIO"] == str(anio_filtro)]

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

      for col in ["Retardo", "Falla"]:
        if col not in resumen.columns:
          resumen[col] = 0

      resumen["TOTAL_INCIDENCIAS"] = resumen["Retardo"] + resumen["Falla"]
      resumen = resumen.sort_values(
          by="TOTAL_INCIDENCIAS", ascending=False
      ).reset_index(drop=True)

      st.dataframe(resumen, use_container_width=True)

      import io

      output = io.BytesIO()
      with pd.ExcelWriter(output, engine="openpyxl") as writer:
        resumen.to_excel(
            writer, index=False, sheet_name="Resumen_General"
        )
        df_filtrado.to_excel(
            writer, index=False, sheet_name="Detalle_Completo"
        )

        for curso_nombre in sorted(df_filtrado["CURSO"].unique()):
          df_curso = df_filtrado[df_filtrado["CURSO"] == curso_nombre]
          resumen_curso = (
              df_curso.pivot_table(
                  index=["CURSO", "NOMBRE_ESTUDIANTE"],
                  columns="TIPO",
                  values="FECHA",
                  aggfunc="count",
                  fill_value=0,
              )
              .reset_index()
          )
          for col in ["Retardo", "Falla"]:
            if col not in resumen_curso.columns:
              resumen_curso[col] = 0
          resumen_curso["TOTAL_INCIDENCIAS"] = (
              resumen_curso["Retardo"] + resumen_curso["Falla"]
          )
          resumen_curso = resumen_curso.sort_values(
              by="TOTAL_INCIDENCIAS", ascending=False
          ).reset_index(drop=True)

          sheet_name_clean = str(curso_nombre).replace("/", "-")[:31]
          resumen_curso.to_excel(
              writer, index=False, sheet_name=sheet_name_clean
          )

      excel_data = output.getvalue()

      st.download_button(
          label="📥 Descargar Reporte en Excel (General + Pestañas por Curso)",
          data=excel_data,
          file_name=f"reporte_asistencia_{anio_filtro}.xlsx",
          mime=(
              "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          ),
          use_container_width=True,
      )

      with st.expander("Ver Historial Detallado"):
        st.dataframe(df_filtrado, use_container_width=True)
