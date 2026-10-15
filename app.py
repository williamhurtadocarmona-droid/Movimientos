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
    "Sube tus extractos bancarios en PDF. La aplicación extraerá todos los"
    " movimientos detallados, los organizará por destinatario y te permitirá"
    " descargar el Excel."
)

uploaded_files = st.file_uploader(
    "Sube tus archivos PDF", type=["pdf"], accept_multiple_files=True
)

if uploaded_files:
  all_transactions = []

  with st.spinner("Procesando y extrayendo todos los movimientos..."):
    for uploaded_file in uploaded_files:
      try:
        reader = pypdf.PdfReader(uploaded_file)

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
        date_pattern = re.compile(r"^(\d{1,2}/\d{2}/\d{4})$")
        money_pattern = re.compile(r"(-?[\d\.]+,\d{2})")

        i = 0
        while i < len(lines):
          line = lines[i].strip()
          if date_pattern.match(line):
            fecha = line
            block_lines = []
            i += 1
            # Recolectar las líneas asociadas a esta transacción hasta la siguiente fecha o pie de página
            while i < len(lines):
              next_line = lines[i].strip()
              if (
                  date_pattern.match(next_line)
                  or "Ponemos a tu disposición" in next_line
                  or "SUPERINTENDENCIA" in next_line
                  or "DETALLE DE CUENTA" in next_line
                  or "Pág" in next_line
              ):
                break
              block_lines.append(next_line)
              i += 1

            block_text = " ".join(block_lines)
            money_matches = money_pattern.findall(block_text)

            if money_matches:
              # En el formato Davivienda, el monto de la transacción es el penúltimo valor monetario (el último es el saldo)
              if len(money_matches) >= 2:
                monto_str = money_matches[-2]
              else:
                monto_str = money_matches[-1]

              cleaned = monto_str.replace(".", "").replace(",", ".")
              try:
                monto = float(cleaned)
                desc = block_text
                for mm in money_matches:
                  desc = desc.replace(mm, "")
                desc = re.sub(r"\s+", " ", desc).strip()

                if not desc or len(desc) < 2:
                  desc = "TRANSACCION BANCARIA"

                tipo = "Egreso" if monto < 0 else "Ingreso"

                all_transactions.append({
                    "Archivo": uploaded_file.name,
                    "Fecha": fecha,
                    "Destinatario / Descripcion": desc,
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
        "No se detectaron transacciones válidas. Revisa que el archivo PDF"
        " corresponda al extracto esperado."
    )
