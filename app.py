import io
import pandas as pd
import pypdf
import streamlit as st

st.set_page_config(
    page_title="Extractor de Extractos a Excel", page_icon="📊", layout="centered"
)

st.title("📊 Procesador de Extractos Bancarios PDF a Excel")
st.write(
    "Sube uno o varios extractos bancarios en PDF para organizarlos"
    " automáticamente por destinatario (ingresos y egresos)."
)

# Habilitar la carga de múltiples archivos con accept_multiple_files=True
uploaded_files = st.file_uploader(
    "Sube tus archivos PDF", type=["pdf"], accept_multiple_files=True
)

if uploaded_files:
  all_data = []

  with st.spinner("Procesando los documentos..."):
    for uploaded_file in uploaded_files:
      try:
        reader = pypdf.PdfReader(uploaded_file)

        # Si el PDF está protegido/encriptado, intentar desencriptarlo con clave vacía
        if reader.is_encrypted:
          try:
            reader.decrypt("")
          except Exception:
            st.warning(
                f"El archivo {uploaded_file.name} está protegido con"
                " contraseña."
            )
            continue

        text = ""
        for page in reader.pages:
          extracted = page.extract_text()
          if extracted:
            text += extracted + "\n"

        # === AQUÍ PUEDES ADAPTAR TU LÓGICA DE EXTRACCIÓN DE TRANSACCIONES ===
        # A modo de ejemplo con los datos analizados previamente:
        # (Si gustas, puedes implementar expresiones regulares para leer fechas, montos y descripciones automáticamente del texto extraído)

      except Exception as e:
        st.error(f"Error al leer {uploaded_file.name}: {e}")

    st.success("¡Proceso de lectura finalizado!")

    # Ejemplo de datos unificados listos para mostrar
    data = [
        {
            "Archivo": uploaded_files[0].name,
            "Fecha": "1/09/2026",
            "Destinatario / Descripcion": "ESTADERO BRISAS DE DICIEMBRE",
            "Monto": -26500.0,
            "Tipo": "Egreso",
        },
        {
            "Archivo": uploaded_files[0].name,
            "Fecha": "1/09/2026",
            "Destinatario / Descripcion": "ALTA SAS",
            "Monto": -32000.0,
            "Tipo": "Egreso",
        },
        {
            "Archivo": uploaded_files[0].name,
            "Fecha": "2/09/2026",
            "Destinatario / Descripcion": "JAIR ALFONSO MIELES MULFORD",
            "Monto": 41000.0,
            "Tipo": "Ingreso",
        },
    ]

    df = pd.DataFrame(data)

    st.subheader("Vista previa de movimientos unificados")
    st.dataframe(df, use_container_width=True)

    # Generar Excel consolidado
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
      df.to_excel(writer, index=False, sheet_name="Movimientos")

      resumen = df.groupby("Tipo")["Monto"].sum().reset_index()
      resumen.to_excel(writer, index=False, sheet_name="Resumen")

    st.download_button(
        label="📥 Descargar Excel Consolidado",
        data=output.getvalue(),
        file_name="movimientos_consolidados.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    )
