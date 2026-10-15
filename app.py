import io
import re
import pandas as pd
from pdfminer.high_level import extract_text
import streamlit as st

st.set_page_config(
    page_title="Extractor de Extractos a Excel", page_icon="📊", layout="wide"
)

st.title("📊 Procesador de Extractos Bancarios Davivienda")
st.write(
    "Sube tus extractos en PDF. La aplicación unificará la línea de la"
    " transacción con el nombre comercial correspondiente y generará tu Excel"
    " organizado."
)

uploaded_files = st.file_uploader(
    "Sube tus archivos PDF", type=["pdf"], accept_multiple_files=True
)

if uploaded_files:
  all_transactions = []

  with st.spinner("Procesando y unificando filas comerciales..."):
    for uploaded_file in uploaded_files:
      try:
        text = extract_text(uploaded_file)
        lines = [line.strip() for line in text.split("\n") if line.strip()]

        date_pattern = re.compile(r"^(\d{1,2}/\d{2}/\d{4})$")
        money_pattern = re.compile(r"(-?[\d\.]+,\d{2})")

        i = 0
        while i < len(lines):
          line = lines[i]

          # Si encontramos una línea con fecha, es el inicio de una transacción
          if date_pattern.match(line):
            fecha = line
            transac_lines = [line]
            i += 1

            # Recolectar elementos hasta la siguiente fecha o sección del documento
            while i < len(lines):
              next_line = lines[i]
              if (
                  date_pattern.match(next_line)
                  or "Ponemos a tu disposición" in next_line
                  or "SUPERINTENDENCIA" in next_line
                  or "DETALLE DE CUENTA" in next_line
                  or "Pág" in next_line
                  or "RESUMEN CUENTA" in next_line
              ):
                break
              transac_lines.append(next_line)
              i += 1

            # Unir el bloque de texto de esta transacción
            block_text = " " .join(transac_lines)
            money_matches = money_pattern.findall(block_text)

            if money_matches:
              # En Davivienda el monto suele ser el penúltimo valor monetario (el último es el saldo)
              if len(money_matches) >= 2:
                monto_str = money_matches[-2]
              else:
                monto_str = money_matches[-1]

              cleaned_monto = monto_str.replace(".", "").replace(",", ".")
              try:
                monto = float(cleaned_monto)

                # Construir una descripción limpia combinando el código BREB/transacción y el nombre comercial de la línea siguiente
                # Buscamos líneas que no sean fechas ni montos ni oficinas genéricas
                desc_parts = []
                for tl in transac_lines[
                    1:
                ]:  # Omitir la fecha de la primera posición
                  if (
                      not money_pattern.search(tl)
                      and "CENTRAL DE C" not in tl
                      and "BARRANQUILLA" not in tl
                  ):
                    desc_parts.append(tl)

                descripcion = (
                    " - ".join(desc_parts)
                    if desc_parts
                    else "TRANSACCION BANCARIA"
                )
                descripcion = re.sub(r"\s+", " ", descripcion).strip()

                tipo = "Egreso" if monto < 0 else "Ingreso"

                all_transactions.append({
                    "Archivo": uploaded_file.name,
                    "Fecha": fecha,
                    "Destinatario / Descripcion": descripcion,
                    "Monto": monto,
                    "Tipo": tipo,
                })
              except ValueError:
                continue
          else:
            i += 1

      except Exception as e:
        st.error(f"Error procesando {uploaded_file.name}: {e}")

  if all_transactions:
    df = pd.DataFrame(all_transactions)

    st.success(
        f"¡Se han extraído correctamente {len(df)} movimientos con su respectivo"
        " nombre comercial!"
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
        "No se detectaron transacciones. Comprueba que el archivo PDF sea el"
        " extracto correcto."
    )
