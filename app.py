import io
import re
from pdfminer.high_level import extract_text
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Procesador Consolidado y Neto", page_icon="📊", layout="wide"
)

st.title("📊 Procesador de Ingresos, Egresos y Consolidado Neto")
st.write(
    "Sube tus archivos en formato **PDF**, **Excel** o **CSV**. La aplicación"
    " unificará los movimientos y calculará el **balance neto resultante** por"
    " cada destinatario/concepto para evitar repeticiones."
)

uploaded_files = st.file_uploader(
    "Sube tus archivos",
    type=["pdf", "xlsx", "xls", "csv"],
    accept_multiple_files=True,
)

if uploaded_files:
  all_transactions = []

  with st.spinner("Procesando y agrupando transacciones..."):
    for uploaded_file in uploaded_files:
      file_name = uploaded_file.name
      file_extension = file_name.split(".")[-1].lower()

      try:
        # 1. PROCESAMIENTO DE ARCHIVOS PDF (Extractos Davivienda)
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
                if len(money_matches) >= 2:
                  monto_str = money_matches[-2]
                else:
                  monto_str = money_matches[-1]

                cleaned_monto = monto_str.replace(".", "").replace(",", ".")
                try:
                  monto = float(cleaned_monto)
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
              fecha = str(row[col_fecha]).split(" ")[0]
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
        f"¡Se han procesado exitosamente {len(df_final)} movimientos en total!"
    )

    # Crear la vista agrupada con valores netos resultantes
    df_neto = (
        df_final.groupby("Destinatario / Descripcion")
        .agg(
            Cantidad_Movimientos=("Monto", "count"),
            Monto_Neto_Resultante=("Monto", "sum"),
        )
        .reset_index()
    )

    df_neto["Tipo Neto"] = df_neto["Monto_Neto_Resultante"].apply(
        lambda x: "Ingreso Neto" if x >= 0 else "Egreso Neto"
    )
    # Reordenar columnas
    df_neto = df_neto[[
        "Destinatario / Descripcion",
        "Cantidad_Movimientos",
        "Monto_Neto_Resultante",
        "Tipo Neto",
    ]]

    st.subheader("Vista consolidada por Destinatario (Valores Netos)")
    st.dataframe(df_neto, use_container_width=True)

    # Generar Excel estructurado con múltiples pestañas profesionales
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
      # Pestaña 1: Todos los movimientos detallados
      df_final.to_excel(writer, index=False, sheet_name="Movimientos Detallados")

      # Pestaña 2: Consolidado neto sin repeticiones
      df_neto.to_excel(
          writer, index=False, sheet_name="Consolidado Neto por Destinatario"
      )

    st.download_button(
        label="📥 Descargar Excel con Consolidado Neto",
        data=output.getvalue(),
        file_name="reporte_neto_ingresos_egresos.xlsx",
        mime=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        ),
    )
  else:
    st.warning("No se encontraron transacciones válidas en los archivos subidos.")
