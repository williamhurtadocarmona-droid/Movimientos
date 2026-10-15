import pandas as pd
import streamlit as st


@st.cache_data
def consolidar_archivos(file_nequi, file_movs):
  # 1. Leer documento.xlsx (Nequi) detectando dinámicamente la fila de cabecera
  df1_raw = pd.read_excel(file_nequi, sheet_name='Table 1')
  header_idx = 7
  for idx, row in df1_raw.iterrows():
    if row.astype(str).str.contains('Fecha del movimiento').any():
      header_idx = idx
      break

  df1 = pd.read_excel(
      file_nequi, sheet_name='Table 1', skiprows=header_idx + 1
  )
  df1 = df1.iloc[:, [0, 1, 3]]
  df1.columns = ['Fecha', 'Descripción', 'Valor']

  # 2. Leer el segundo archivo de movimientos (.xls)
  df2 = pd.read_excel(file_movs, sheet_name='Movimientos')
  val_debito = (
      df2['Valor Débito']
      if 'Valor Débito' in df2.columns
      else df2.iloc[:, 2]
  )
  val_credito = (
      df2['Valor Crédito']
      if 'Valor Crédito' in df2.columns
      else df2.iloc[:, 3]
  )
  df2['Valor'] = val_debito.fillna(0) + val_credito.fillna(0)
  df2 = df2[['Fecha', 'Descripción', 'Valor']]

  # 3. Unir ambos DataFrames
  df_total = pd.concat([df1, df2], ignore_index=True)
  df_total['Valor'] = pd.to_numeric(df_total['Valor'], errors='coerce').fillna(
      0
  )

  # 4. Filtrar solo salidas de dinero (valores negativos)
  df_salidas = df_total[df_total['Valor'] < 0].copy()

  # 5. Limpiar descripciones para extraer el nombre de la persona
  limpiar_textos = (
      df_salidas['Descripción']
      .astype(str)
      .str.replace('Para ', '', case=False)
      .str.replace('ENVIO CON BRE-B A: ', '', case=False)
      .str.strip()
  )
  df_salidas['Persona'] = limpiar_textos
  df_salidas['Valor_Absoluto'] = df_salidas['Valor'].abs()

  # 6. Agrupar por Persona: sumar los valores y contar los movimientos
  df_consolidado = (
      df_salidas.groupby('Persona')
      .agg(
          Total_Enviado=('Valor_Absoluto', 'sum'),
          Cantidad_Envios=('Valor_Absoluto', 'count'),
      )
      .reset_index()
  )

  # Ordenar de mayor a menor monto
  df_consolidado = df_consolidado.sort_values(
      by='Total_Enviado', ascending=False
  ).reset_index(drop=True)

  return df_consolidado


# --- Interfaz en Streamlit ---
st.title('Consolidador de Movimientos por Persona')

uploaded_nequi = st.file_uploader('Sube documento.xlsx (Nequi)', type=['xlsx'])
uploaded_movs = st.file_uploader(
    'Sube Movs_CA_15092026_15102026.xls', type=['xls', 'xlsx']
)

if uploaded_nequi and uploaded_movs:
  df_resultado = consolidar_archivos(uploaded_nequi, uploaded_movs)

  st.success('¡Consolidado generado con éxito!')

  # Mostramos la tabla con las columnas separadas
  st.dataframe(df_resultado, use_container_width=True)

  # Botón de descarga en CSV
  csv = df_resultado.to_csv(index=False).encode('utf-8')
  st.download_button(
      label='Descargar consolidado en CSV',
      data=csv,
      file_name='movimientos_por_persona.csv',
      mime='text/csv',
  )
