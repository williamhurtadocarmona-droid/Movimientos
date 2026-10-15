import io
import re
import pandas as pd
import pypdf
import streamlit as st

st.set_page_config(
    page_title="Extractor de Extractos a Excel", page_icon="📊", layout="wide"
)

st.title("📊 Procesador Multiartefacto de Extractos Bancarios PDF")
st.write(
    "Sube uno o varios extractos bancarios en PDF. La aplicación unificará"
    " todos los movimientos de todos los archivos cargados."
)

uploaded_files = st.file_uploader(
    "Sube tus archivos PDF", type=["pdf"], accept_multiple_files=True
)

if uploaded_files:
  all_transactions = []

  with st.spinner("Procesando y consolidando los archivos PDF..."):
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
        date_pattern = re.compile(r"^(\d{1,2}/\d{2}/\d{4})")

        for line in lines:
          line_str = line.strip()
          match = date_pattern.match(line_str)
          if match:
            fecha = match.group(1)
            # Buscar patrones monetarios (ej. -26.500,00 o 4.682.542,00)
            monto_match = re.findall(r"(-?[\d\.]+,\d{2})", line_str)
            if monto_match:
              monto_str = (
                  monto_match[-1].replace(".", "").replace(",", ".")
              )
              try:
                monto = float(monto_str)
                descripcion = (
                    line_str.replace(fecha, "")
                    .replace(monto_match[-1], "")
                    .strip()
                )
                if not descripcion:
                  descripcion = "TRANSACCION BANCARIA"

                tipo = "Egreso" if monto < 0 else "Ingreso"

                # Acumular correctamente en la lista global
                all_transactions.append({
                    "Archivo": uploaded_file.name,
                    "Fecha": fecha,
                    "Destinatario / Descripcion": descripcion,
                    "Monto": monto,
                    "Tipo": tipo,
                })
              except ValueError:
                continue

      except Exception as e:
        st.error(f"Error procesando {uploaded_file.name}: {e}")

  if all_transactions:
    df = pd.DataFrame(all_transactions)

    st.success(
        f"¡Se han consolidado un total de {len(df)} movimientos de"
        f" {len(uploaded_files)} archivo(s)!"
    )

    st.subheader("Vista previa de movimientos consolidados")
    st.dataframe(df, use_container_width=True)

    # Generar Excel con hojas de Movimientos y Resumen
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
      df.to_excel(writer, index=False, sheet_name="Movimientos Consolidados")

      resumen = (
          df.groupby(["Archivo", "Tipo"])["Monto"]
          .agg(["count", "sum"])
          .reset_index()
      )
      resumen.columns = ["Archivo", "Tipo", "Cantidad", "Total Monto"]
      resumen.to_excel(writer, index=False, sheet_name="Resumen por Archivo")

    st.download_button(
        label="📥 Descargar Excel Consolidado General",
        data=output.getvalue(),
        file_name="movimientos_consolidados_general.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    )
  else:
    st.warning(
        "No se detectaron transacciones válidas con el formato esperado en los"
        " archivos subidos."
    )
