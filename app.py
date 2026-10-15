import io
import re
from pdfminer.high_level import extract_text
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Procesador de Movimientos", page_icon="📊", layout="wide"
)

st.title("📊 Procesador de Ingresos y Egresos (PDF y Excel)")
st.write(
    "Sube tus archivos en formato **PDF**, **Excel (.xls, .xlsx)** o **CSV** para"
    " organizarlos automáticamente por destinatario, ingresos y egresos."
)

uploaded_files = st.file_uploader(
    "Sube tus archivos",
    type=["pdf", "xlsx", "xls", "csv"],
    accept_multiple_files=True,
)

if uploaded_files:
  all_transactions = []

  with st.spinner("Procesando y organizando los documentos..."):
    for uploaded_file in uploaded_files:
      file_name = uploaded_file.name
      file_extension = file_name.split(".")[-1].lower()

      try:
        # 1. Procesamiento si es PDF
        if file_extension == "pdf":
          text = extract_text(uploaded_file)
          lines = [line.strip() for line in text.split("\n") if line.strip()]

          dates = []
          amounts = []
          descriptions = []

          date_pattern = re.compile(r"^(\d{1,2}/\d{2}/\d{4})$")
          money_pattern = re.compile(r"(-?[\d\.]+,\d{2})")

          for line in lines:
            if date_pattern.match(line):
              dates.append(line)
            elif money_pattern.match(line):
              amounts.append(line)
            elif any(
                k in line.upper()
                for k in [
                    "BREB",
                    "COMPRA",
                    "RETIRO",
                    "ESTADERO",
                    "ALTA",
                    "ABRAHAM",
                    "SUBLIMEX",
                    "MELISSA",
                    "COMISION",
                    "JAIR",
                    "PAGO",
                    "NEQUI",
                    "DAVIBANK",
                    "ACH",
                ]
            ):
              descriptions.append(line)

          min_len = min(len(dates), len(amounts))
          for idx in range(min_len):
            fecha = dates[idx]
            monto_str = amounts[idx]
            cleaned_monto = (
                monto_str.replace(".", "").replace(",", ".").replace(" ", "")
            )
            try:
              monto = float(cleaned_monto)
              desc = (
                  descriptions[idx]
                  if idx < len(descriptions)
                  else "TRANSACCION BANCARIA"
              )
              tipo = "Egreso" if monto < 0 else "Ingreso"

              all_transactions.append({
                  "Archivo": file_name,
                  "Fecha": fecha,
                  "Destinatario / Descripcion": desc,
                  "Monto": monto,
                  "Tipo": tipo,
              })
            except ValueError:
              continue

        # 2. Procesamiento si es Excel (.xlsx, .xls)
        elif file_extension in ["xlsx", "xls"]:
          df_excel = pd.read_excel(uploaded_file, sheet_name=0)
          # Limpiar nombres de columnas
          df_excel.columns = [str(c).strip().lower() for c in df_excel.columns]

          # Buscar automáticamente las columnas clave
          col_fecha = next(
              (c for c in df_excel.columns if "fecha" in c), df_excel.columns[0]
          )
          col_desc = next(
              (
                  c
                  for c in df_excel.columns
                  if "desc" in c or "dest" in c or "concepto" in c or "detalle" in c
              ),
              (
                  df_excel.columns[1]
                  if len(df_excel.columns) > 1
                  else df_excel.columns[0]
              ),
          )
          col_monto = next(
              (
                  c
                  for c in df_excel.columns
                  if "monto" in c or "valor" in c or "neto" in c or "debito" in c
              ),
              df_excel.columns[-1],
          )

          for _, row in df_excel.iterrows():
            try:
              fecha = str(row[col_fecha])
              desc = str(row[col_desc])
              monto_val = str(row[col_monto])
              # Limpiar formato de moneda
              cleaned_monto = (
                  monto_val.replace("$", "")
                  .replace(".", "")
                  .replace(",", ".")
                  .strip()
              )
              monto = float(cleaned_monto)
              tipo = "Egreso" if monto < 0 else "Ingreso"

              all_transactions.append({
                  "Archivo": file_name,
                  "Fecha": fecha,
                  "Destinatario / Descripcion": desc,
                  "Monto": monto,
                  "Tipo": tipo,
              })
            except Exception:
              continue

        # 3. Procesamiento si es CSV
        elif file_extension == "csv":
          df_csv = pd.read_csv(uploaded_file)
          df_csv.columns = [str(c).strip().lower() for c in df_csv.columns]

          col_fecha = next(
              (c for c in df_csv.columns if "fecha" in c), df_csv.columns[0]
          )
          col_desc = next(
              (
                  c
                  for c in df_csv.columns
                  if "desc" in c or "dest" in c or "concepto" in c
              ),
              (
                  df_csv.columns[1]
                  if len(df_csv.columns) > 1
                  else df_csv.columns[0]
              ),
          )
          col_monto = next(
              (
                  c
                  for c in df_csv.columns
                  if "monto" in c or "valor" in c or "neto" in c
              ),
              df_csv.columns[-1],
          )

          for _, row in df_csv.iterrows():
            try:
              fecha = str(row[col_fecha])
              desc = str(row[col_desc])
              cleaned_monto = (
                  str(row[col_monto])
                  .replace("$", "")
                  .replace(".", "")
                  .replace(",", ".")
                  .strip()
              )
              monto = float(cleaned_monto)
              tipo = "Egreso" if monto < 0 else "Ingreso"

              all_transactions.append({
                  "Archivo": file_name,
                  "Fecha": fecha,
                  "Destinatario / Descripcion": desc,
                  "Monto": monto,
                  "Tipo": tipo,
              })
            except Exception:
              continue

      except Exception as e:
        st.error(f"Error procesando {file_name}: {e}")

  if all_transactions:
    df_final = pd.DataFrame(all_transactions)

    st.success(
        f"¡Se han procesado exitosamente {len(df_final)} movimientos en total!"
    )
    st.subheader("Vista previa unificada")
    st.dataframe(df_final, use_container_width=True)

    # Generar Excel estructurado
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
      df_final.to_excel(writer, index=False, sheet_name="Movimientos")

      resumen = (
          df_final.groupby(["Tipo", "Destinatario / Descripcion"])["Monto"]
          .agg(["count", "sum"])
          .reset_index()
      )
      resumen.columns = [
          "Tipo",
          "Destinatario / Descripcion",
          "Cantidad",
          "Total Monto",
      ]
      resumen.to_excel(
          writer, index=False, sheet_name="Resumen por Destinatario"
      )

    st.download_button(
        label="📥 Descargar Excel Organizado",
        data=output.getvalue(),
        file_name="reporte_ingresos_egresos.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    )
  else:
    st.warning("No se encontraron transacciones válidas en los archivos subidos.")
