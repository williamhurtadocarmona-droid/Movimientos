import io
import re
import pandas as pd
import pypdf
import streamlit as st

st.set_page_config(
    page_title="Extractor de Extractos a Excel", page_icon="📊", layout="wide"
)

st.title("📊 Procesador Automático de Extractos Bancarios PDF a Excel")
st.write(
    "Sube tus extractos en PDF. La aplicación extraerá todas las"
    " transacciones, las organizará por destinatario y generará un Excel"
    " consolidado."
)

uploaded_files = st.file_uploader(
    "Sube tus archivos PDF", type=["pdf"], accept_multiple_files=True
)

if uploaded_files:
  all_transactions = []

  with st.spinner("Procesando y extrayendo movimientos de los PDFs..."):
    for uploaded_file in uploaded_files:
      try:
        reader = pypdf.PdfReader(uploaded_file)

        # Desencriptar si es necesario
        if reader.is_encrypted:
          try:
            reader.decrypt("")
          except Exception:
            st.warning(
                f"El archivo {uploaded_file.name} está protegido con"
                " contraseña."
            )
            continue

        full_text = ""
        for page in reader.pages:
          extracted = page.extract_text()
          if extracted:
            full_text += extracted + "\n"

        # Lógica de extracción de transacciones basada en patrones de extractos bancarios (ej. formato Davivienda)
        # Buscamos líneas que contengan fechas (DD/MM/YYYY) y montos con formato numérico
        lines = full_text.split("\n")

        # Patrón para detectar fechas al inicio de línea
        date_pattern = re.compile(r"^(\d{1,2}/\d{2}/\d{4})")

        for line in lines:
          match = date_pattern.match(line.strip())
          if match:
            fecha = match.group(1)
            # Limpiar la línea para aislar montos y descripciones
            # Buscamos montos al final de la línea (ej. -26.500,00 o 4.000,00)
            monto_match = re.findall(
                r"(-?[\d\.]+,\d{2})", line
            )  # Busca números con formato colombiano/español
            if monto_match:
              # Tomamos el último valor encontrado en la línea como el monto
              monto_str = (
                  monto_match[-1].replace(".", "").replace(",", ".")
              )  # Convertir a formato float de python
              try:
                monto = float(monto_str)

                # La descripción es todo lo que está entre la fecha y el monto
                descripcion = line.replace(fecha, "").replace(monto_match[-1], "").strip()
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
              except ValueError:
                continue

      except Exception as e:
        st.error(f"Error procesando {uploaded_file.name}: {e}")

  if all_transactions:
    df = pd.DataFrame(all_transactions)

    st.success(
        f"¡Se han extraído exitosamente {len(df)} movimientos de los archivos"
        " cargados!"
    )

    st.subheader("Vista previa de todos los movimientos detectados")
    st.dataframe(df, use_container_width=True)

    # Generar Excel consolidado profesional
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
      df.to_excel(writer, index=False, sheet_name="Movimientos")

      # Hoja de resumen por tipo
      resumen = df.groupby("Tipo")["Monto"].agg(["count", "sum"]).reset_index()
      resumen.columns = ["Tipo", "Cantidad", "Total Monto"]
      resumen.to_excel(writer, index=False, sheet_name="Resumen")

    st.download_button(
        label="📥 Descargar Excel Consolidado Completo",
        data=output.getvalue(),
        file_name="movimientos_completos_organizados.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    )
  else:
    st.warning(
        "No se pudieron extraer transacciones automáticas. Asegúrate de que el"
        " formato del PDF coincida con el estándar esperado."
    )
