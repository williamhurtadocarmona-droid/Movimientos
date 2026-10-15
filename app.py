import io
import pandas as pd
import pdfplumber
import streamlit as st

st.set_page_config(
    page_title="Extractor de Extractos a Excel", page_icon="📊", layout="wide"
)

st.title("📊 Procesador Automático de Extractos Bancarios PDF")
st.write(
    "Sube tus extractos bancarios en PDF. La aplicación extraerá todos los"
    " movimientos organizados por fecha, descripción, ingresos y egresos."
)

uploaded_files = st.file_uploader(
    "Sube tus archivos PDF", type=["pdf"], accept_multiple_files=True
)

if uploaded_files:
  all_transactions = []

  with st.spinner("Extrayendo tablas y movimientos con precisión..."):
    for uploaded_file in uploaded_files:
      try:
        with pdfplumber.open(uploaded_file) as pdf:
          for page in pdf.pages:
            # Extraer tablas de la página
            tables = page.extract_tables()
            for table in tables:
              for row in table:
                # Filtrar filas que contengan datos de transacciones (buscando formato de fecha)
                row_text = " ".join([str(cell) for cell in row if cell])
                if "/" in row_text and (
                    "-" in row_text or "," in row_text or "." in row_text
                ):
                  # Intentar identificar fecha, descripción y monto dentro de las celdas de la tabla
                  cells = [c.strip() for c in row if c and c.strip()]
                  if len(cells) >= 3:
                    fecha = cells[0]
                    # Buscar la celda que tenga formato de valor monetario
                    for cell in reversed(cells):
                      cleaned_cell = (
                          cell.replace(".", "").replace(",", ".").replace("$", "")
                      )
                      try:
                        monto = float(cleaned_cell)
                        # Si encontramos el monto, la descripción es el texto intermedio
                        descripcion = " ".join(
                            [
                                c
                                for c in cells
                                if c != fecha and c != cell and not "-" in c
                            ]
                        )
                        if not descripcion:
                          descripcion = "TRANSACCION BANCARIA"

                        tipo = "Egreso" if monto < 0 else "Ingreso"

                        all_transactions.append({
                            "Archivo": uploaded_file.name,
                            "Fecha": fecha,
                            "Destinatario / Descripcion": descripcion,
                            "Monto": monto,
                            "Tipo": tipo,
                        })
                        break
                      except ValueError:
                        continue

      except Exception as e:
        st.error(f"Error procesando {uploaded_file.name}: {e}")

  if all_transactions:
    df = pd.DataFrame(all_transactions)

    st.success(
        f"¡Se han extraído correctamente {len(df)} movimientos detallados de"
        " los archivos cargados!"
    )

    st.subheader("Vista previa de movimientos organizados")
    st.dataframe(df, use_container_width=True)

    # Generar Excel estructurado
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
      df.to_excel(writer, index=False, sheet_name="Movimientos")

      resumen = (
          df.groupby(["Archivo", "Tipo"])["Monto"]
          .agg(["count", "sum"])
          .reset_index()
      )
      resumen.columns = ["Archivo", "Tipo", "Cantidad", "Total Monto"]
      resumen.to_excel(writer, index=False, sheet_name="Resumen")

    st.download_button(
        label="📥 Descargar Excel Organizado",
        data=output.getvalue(),
        file_name="movimientos_banco_organizados.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    )
  else:
    st.warning(
        "No se detectaron transacciones en formato de tabla. Asegúrate de que el"
        " archivo sea el extracto bancario en PDF."
    )
