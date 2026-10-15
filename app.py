import io
import re
import pandas as pd
import pypdf
import streamlit as st

st.set_page_config(
    page_title="Extractor de Extractos a Excel", page_icon="📊", layout="wide"
)

st.title("📊 Procesador Automático de Extractos Bancarios PDF")
st.write(
    "Sube tus extractos bancarios en PDF (como los de Davivienda). La"
    " aplicación extraerá todos los movimientos, los clasificará y te"
    " permitirá descargar el Excel organizado."
)

uploaded_files = st.file_uploader(
    "Sube tus archivos PDF", type=["pdf"], accept_multiple_files=True
)

if uploaded_files:
  all_transactions = []

  with st.spinner("Analizando y extrayendo movimientos de los PDFs..."):
    for uploaded_file in uploaded_files:
      try:
        reader = pypdf.PdfReader(uploaded_file)

        # Desencriptar si el PDF está protegido
        if reader.is_encrypted:
          try:
            reader.decrypt("")
          except Exception:
            st.warning(
                f"El archivo {uploaded_file.name} está protegido y no se pudo"
                " leer."
            )
            continue

        full_text = ""
        for page in reader.pages:
          extracted = page.extract_text()
          if extracted:
            full_text += extracted + "\n"

        lines = full_text.split("\n")
        date_regex = re.compile(r"^(\d{1,2}/\d{2}/\d{4})$")

        # Parser robusto línea por línea adaptado a tablas bancarias
        i = 0
        while i < len(lines):
          line = lines[i].strip()
          if date_regex.match(line):
            fecha = line
            desc_lines = []
            i += 1
            # Recolectar líneas de descripción hasta la siguiente fecha o sección del pie de página
            while i < len(lines):
              next_line = lines[i].strip()
              if (
                  date_regex.match(next_line)
                  or "Ponemos a tu disposición" in next_line
                  or "SUPERINTENDENCIA" in next_line
                  or "DETALLE DE CUENTA" in next_line
              ):
                break
              desc_lines.append(next_line)
              i += 1

            combined_text = " ".join(desc_lines)
            # Buscar montos monetarios en formato colombiano (ej. -26.500,00 o 41.000,00)
            money_matches = re.findall(r"(-?[\d\.]+,\d{2})", combined_text)

            if money_matches:
              # En extractos de este tipo, el monto de la transacción suele ser el penúltimo valor monetario (antes del saldo)
              if len(money_matches) >= 2:
                monto_str = money_matches[-2]
              else:
                monto_str = money_matches[-1]

              cleaned_monto = monto_str.replace(".", "").replace(",", ".")
              try:
                monto = float(cleaned_monto)

                # Extraer la descripción eliminando los montos detectados
                description = combined_text
                for mm in money_matches:
                  description = description.replace(mm, "")
                description = re.sub(r"\s+", " ", description).strip()

                if not description:
                  description = "TRANSACCION BANCARIA"

                tipo = "Egreso" if monto < 0 else "Ingreso"

                all_transactions.append({
                    "Archivo": uploaded_file.name,
                    "Fecha": fecha,
                    "Destinatario / Descripcion": description,
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
        f"¡Se han extraído correctamente {len(df)} movimientos de los archivos"
        " cargados!"
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
        "No se detectaron transacciones. Revisa que el PDF sea un extracto"
        " bancario válido."
    )
