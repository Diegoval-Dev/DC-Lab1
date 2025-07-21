import pandas as pd

# Ruta del archivo
file_path = './Estadisticas_historicas_comercializacion.xlsx'

# Cargar hojas IMPORTACION y CONSUMO con encabezado en la fila 7 (índice 6)
sheets = pd.read_excel(
    file_path,
    sheet_name=['IMPORTACION', 'CONSUMO'],
    header=6
)

# Limpiar columnas vacías y estandarizar nombres
df_imp = sheets['IMPORTACION'].dropna(axis=1, how='all')
df_cons = sheets['CONSUMO'].dropna(axis=1, how='all')

df_imp.columns = df_imp.columns.str.strip().str.lower()
df_cons.columns = df_cons.columns.str.strip().str.lower()

# Agregar columna de origen
df_imp['hoja'] = 'IMPORTACION'
df_cons['hoja'] = 'CONSUMO'

# Combinar ambas hojas
df = pd.concat([df_imp, df_cons], ignore_index=True)

# Crear columna diesel
df['diesel'] = (
    df['diesel bajo azufre']
  + df['diesel ultra bajo azufre']
  + df['diesel alto azufre']
)

# Seleccionar variables de interés
variables_objetivo = ['fecha', 'gasolina regular', 'gasolina superior', 'diesel']
df_seleccionado = df[variables_objetivo + ['hoja']]

# Extraer resumen para cada variable
info_codebook = {}
for col in df_seleccionado.columns:
    if col != 'hoja':
        info_codebook[col] = {
            'tipo': str(df_seleccionado[col].dtype),
            'valores únicos': df_seleccionado[col].nunique(),
            'nulos': df_seleccionado[col].isna().sum(),
            'min': df_seleccionado[col].min() if pd.api.types.is_numeric_dtype(df_seleccionado[col]) else None,
            'max': df_seleccionado[col].max() if pd.api.types.is_numeric_dtype(df_seleccionado[col]) else None,
        }

# Mostrar resultado
for var, meta in info_codebook.items():
    print(f'Variable: {var}')
    for k, v in meta.items():
        print(f'  {k}: {v}')
    print()
