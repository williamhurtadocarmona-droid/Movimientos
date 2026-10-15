import io
import re
from pdfminer.high_level import extract_text
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Procesador Unificado de Movimientos",
    page_icon="📊",
    layout="wide",
)

st.title("📊 Procesador de Ingresos y Egresos (PDF, Excel y CSV)")
st.write(
    "Sube tus archivos bancarios en formato **PDF** (extractos Davivienda), así"
    " como tus planillas en **Excel (.xls, .xlsx)** o **CSV**. La aplicación"
    " unificará todos los movimientos organizados por destinatario, ingresos y"
    " egresos."
)

uploaded_files = st.file_uploader(
    "Sube tus archivos",
    type=["pdf", "xlsx", "xls", "csv"],
    accept_multiple_files=True,
)

if uploaded_files:
  all_transactions = []

  with st.spinner("Procesando, alineando y unificando los documentos..."):
    for uploaded_file in uploaded_files:
      file_name = uploaded_file.name
      file_extension = file_name.split(".")[-1].lower()

      try:
        # 1. PROCESAMIENTO DE ARCHIVOS PDF (Estructura de doble línea Davivienda)
        if file_extension == "pdf":
          text = extract_text(uploaded_file)
          lines = [line.strip() for line in text.split("\n") if line.strip()]

          date_pattern = re.compile(r"^(\d{1,2}/\d{2}/\d{4})$")
          money_pattern = re.compile(r"(-?[\d\.]+,\d{2})")

          i = 0
          while i < len(lines):
            line = lines[i]
            if date_pattern.match(line):
              fecha = line
              transac_lines = [line]
              i += 1

              # Recolectar líneas secundarias (como el nombre comercial) hasta la siguiente fecha
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

              block_text = " ".join(transac_lines)
              money_matches = money_pattern.findall(block_text)

              if money_matches:
                # El monto de la transacción suele ser el penúltimo valor monetario (el último es el saldo)
                if len(money_matches) >= 2:
                  monto_str = money_matches[-2]
                else:
                  monto_str = money_matches[-1]

                cleaned_monto = monto_str.replace(".", "").replace(",", ".")
                try:
                  monto = float(cleaned_monto)

                  # Extraer el nombre comercial y código de la transferencia
                  desc_parts = []
                  for tl in transac_lines[1:]:
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
                      "Archivo": file_name,
                      "Fecha": fecha,
                      "Destinatario / Descripcion": descripcion,
                      "Monto": monto,
                      "Tipo": tipo,
                  })
                except ValueError:
                  continue
            else:
              i += 1

        # 2. PROCESAMIENTO DE ARCHIVOS EXCEL (.xlsx, .xls)
        elif file_extension in ["xlsx", "xls"]:
          df_excel = pd.read_excel(uploaded_file, sheet_name=0)
          df_excel.columns = [str(c).strip().lower() for c in df_excel.columns]

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
              fecha = str(row[col_fecha]).split(" ")[
                  0
              ]  # Limpiar formato de fecha si trae hora
              desc = str(row[col_desc])
              monto_val = str(row[col_monto])
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

        # 3. PROCESAMIENTO DE ARCHIVOS CSV
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
              fecha = str(row[col_fecha]).split(" ")[0]
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
        f"¡Se han procesado y unificado exitosamente {len(df_final)} movimientos"
        " en total!"
    )
    st.subheader("Vista previa unificada")
    st.dataframe(df_final, use_container_width=True)

    # Generar Excel estructurado con pestañas profesionales
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
      # Pestaña 1: Todos los movimientos detallados
      df_final.to_excel(writer, index=False, sheet_name="Movimientos Unificados")

      # Pestaña 2: Resumen organizado por Tipo y Destinatario
      resumen = (
          df_final.groupby(["Tipo", "Destinatario / Descripcion"])["Monto"]
          .agg(["count", "sum"])
          .reset_index()
      )
      resumen.columns = [
          "Tipo",
          "Destinatario / Descripcion",
          "Cantidad de Transacciones",
          "Total Monto",
      ]
      resumen.to_excel(
          writer, index=False, sheet_name="Resumen por Destinatario"
      )

    st.download_button(
        label="📥 Descargar Excel Organizado y Consolidado",
        data=output.getvalue(),
        file_name="reporte_consolidado_ingresos_egresos.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    )
  else:
    st.warning(
        "No se encontraron transacciones válidas. Verifica que los archivos"
        " contengan la estructura esperada."
    )
