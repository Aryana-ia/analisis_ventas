"""
Etapas del pipeline de análisis de ventas — v2 (incluye CONSTRUYE ALIANZA).
Código portado del notebook PROanalisispower actualizado.
Cambios respecto al notebook:
  1. Cada celda envuelta en una función; archivo principal parametrizado;
     exit() -> raise SystemExit(1) (plomería, sin efecto en la lógica).
  2. FIX acordado: dos comas faltantes en clientes_a_prefijar (Paso 5):
     - entre 'JUAN ALBERTO REINA OLAYA' y 'CONSTRUCIONES BARINAS'
     - entre 'CONSTRUYE ALIANZA CA' y 'CONSTRUYE ALIANZA., C.A.'
     Sin estas comas, Python concatenaba los strings y esos clientes
     nunca recibían el prefijo '.'.
"""
import pandas as pd  # noqa: F401 (las funciones re-importan igual que el notebook)


def etapa_1_a_4(archivo_principal):
    import pandas as pd

    print("--- INICIO: Ejecutando Pasos 1 a 4 (Lógica de Fechas Definitiva) ---")

    # --- 1. Carga de Datos ---
    print("\n--- PASO 1: Cargando archivo principal ---")
    columnas_como_texto = {
        'DOCUMENTO': str,
        'CODPROD': str,
        'CODVEND': str,
        'RIF': str,
        'FACTAFECT': str
    }
    try:
        df_empresas = pd.read_excel(
            archivo_principal,
            sheet_name='Empresas',
            dtype=columnas_como_texto
        )
        print("✅ Archivo 'Empresas' cargado correctamente.")
    except FileNotFoundError:
        print(f"❌ Error: No se encontró el archivo '{archivo_principal}'. Abortando.")
        raise SystemExit(1)
    
    # --- 2. Limpieza de Espacios ---
    print("\n--- PASO 2: Limpiando espacios innecesarios ---")
    for col in df_empresas.select_dtypes(include=['object']).columns:
        df_empresas[col] = df_empresas[col].str.strip().str.replace(r'\s+', ' ', regex=True)
    print("✅ Espacios limpiados.")

    # --- 3. Formateo de Fecha (VERSIÓN CON FECHA + HORA + MILISEGUNDOS) ---
    print("\n--- PASO 3: Estandarizando la columna de fecha ---")
    if 'FECHA' in df_empresas.columns:
        print("   - Formato de fecha esperado: YYYY-MM-DD HH:MM:SS.mmm (ej: 2026-01-12 16:12:32.000).")
        print("   - Aplicando la conversión correcta...")

        # %Y = año, %m = mes, %d = día
        # %H = hora, %M = minuto, %S = segundo
        # %f = fracción de segundo (milisegundos/microsegundos)
        df_empresas['FECHA'] = pd.to_datetime(
            df_empresas['FECHA'],
            format='%Y-%m-%d %H:%M:%S.%f',
            errors='coerce'
        )

        # Verificación final
        if df_empresas['FECHA'].isnull().any():
            print("   ⚠️ ADVERTENCIA: Quedaron fechas nulas. Algunos valores no cumplen el formato esperado.")
        else:
            print("✅ ¡Excelente! La columna 'FECHA' se convirtió correctamente.")
    else:
        print("⚠️ Advertencia: No se encontró la columna 'FECHA'.")

    # --- 4. Sustitución de Nombres de Empresa ---
    print("\n--- PASO 4: Sustituyendo nombres en la columna 'EMP' ---")
    mapa_empresas = {
        'ACEROS DE VALENCIA, C.A.': 'ACEVAL',
        'EL PUNTO DEL HIERRO, C.A.': 'BARQUISIMETO',
        'FERREACEVAL, C.A.': 'FERREACEVAL',
        'HIERRO EL ROBLE CCS, C.A.': 'CARACAS',
        'HIERRO METALES EL ROBLE, C.A.': 'MARACAY',
        'HIERROS PORTUGUESA, C.A.': 'PORTUGUESA',
        'FERRETERIA PUNTO DEL HIERRO, C.A.': 'BARCELONA',
        'CONSTRUYE ALIANZA, C.A.': 'CONSTALIANZA'
    
    }
    df_empresas['EMP'] = df_empresas['EMP'].replace(mapa_empresas)
    print("✅ Nombres de empresas sustituidos.")

    # --- BLOQUE DE GUARDADO ---
    print("\n--- GUARDADO DE PROGRESO: Creando archivo Excel para los Pasos 1-4 ---")
    nombre_archivo_salida = 'resultado_pasos_1_a_4.xlsx'
    try:
        df_para_guardar = df_empresas.copy()

        if 'FECHA' in df_para_guardar.columns and pd.api.types.is_datetime64_any_dtype(df_para_guardar['FECHA']):
            # Guardar con fecha y hora (sin milisegundos)
            df_para_guardar['FECHA'] = df_para_guardar['FECHA'].dt.strftime('%d/%m/%Y %H:%M:%S')

        print(f"🔄 Guardando en '{nombre_archivo_salida}'...")
        df_para_guardar.to_excel(nombre_archivo_salida, index=False, sheet_name='Pasos_1_a_4_Definitivo')

        print(f"✅ ¡Éxito! El archivo '{nombre_archivo_salida}' ha sido creado con la lógica de fechas correcta.")

    except Exception as e:
        print(f"❌ Ocurrió un error al guardar el archivo: {e}")

    print("\n--- FIN DEL SCRIPT PARA PASOS 1-4 (Definitivo) ---")

def etapa_5():
    import pandas as pd

    print("--- INICIO: Ejecutando Paso 5 (Unificación por archivo externo + ajuste de vendedores) ---")

    # --- CARGA DEL PROGRESO ANTERIOR ---
    print("\n--- CARGANDO PROGRESO: Leyendo 'resultado_pasos_1_a_4.xlsx' ---")
    archivo_anterior = 'resultado_pasos_1_a_4.xlsx'
    try:
        columnas_como_texto = {
            'DOCUMENTO': str,
            'CODPROD': str,
            'CODVEND': str,
            'RIF': str,
            'FACTAFECT': str
        }
        df_empresas = pd.read_excel(
            archivo_anterior,
            dtype=columnas_como_texto
        )
        print(f"✅ Archivo '{archivo_anterior}' cargado correctamente.")
    except FileNotFoundError:
        print(f"❌ Error: No se encontró el archivo '{archivo_anterior}'.")
        print("   Asegúrate de haber ejecutado el script anterior primero.")
        raise SystemExit(1)

    # --- Limpieza ligera de espacios en columnas de texto ---
    for col in df_empresas.select_dtypes(include=['object']).columns:
        df_empresas[col] = df_empresas[col].astype(str).str.strip().str.replace(r'\s+', ' ', regex=True)

    # --- MODIFICACIÓN 1: Unificar NOMCLIE usando CLIENTES_UNICOS.xlsx ---
    print("\n--- PASO 5: Unificando 'NOMCLIE' usando coincidencias de 'RIF' desde 'CLIENTES_UNICOS.xlsx' ---")
    archivo_clientes_unicos = 'CLIENTES_UNICOS.xlsx'

    try:
        df_clientes_unicos = pd.read_excel(archivo_clientes_unicos, dtype={'RIF': str})
        print(f"✅ Archivo '{archivo_clientes_unicos}' cargado correctamente.")
    except FileNotFoundError:
        print(f"❌ Error: No se encontró el archivo '{archivo_clientes_unicos}'.")
        raise SystemExit(1)

    # Validar columnas requeridas
    for col in ['RIF', 'NOMCLIE']:
        if col not in df_clientes_unicos.columns:
            print(f"❌ Error: La columna '{col}' no existe en '{archivo_clientes_unicos}'.")
            raise SystemExit(1)

    # Normalizar
    df_clientes_unicos['RIF'] = df_clientes_unicos['RIF'].astype(str).str.strip().str.replace(r'\s+', ' ', regex=True)
    df_clientes_unicos['NOMCLIE'] = df_clientes_unicos['NOMCLIE'].astype(str).str.strip().str.replace(r'\s+', ' ', regex=True)
    df_empresas['RIF'] = df_empresas['RIF'].astype(str).str.strip().str.replace(r'\s+', ' ', regex=True)

    # Mapa RIF -> NOMCLIE
    mapa_clientes = (
        df_clientes_unicos
        .dropna(subset=['RIF', 'NOMCLIE'])
        .drop_duplicates(subset=['RIF'], keep='first')
        .set_index('RIF')['NOMCLIE']
    )

    # Guardar NOMCLIE original por trazabilidad
    if 'NOMCLIE' not in df_empresas.columns:
        df_empresas['NOMCLIE'] = pd.NA
    df_empresas['NOMCLIE_ORIGINAL'] = df_empresas['NOMCLIE']

    # Detectar coincidencias por RIF
    coincide_rif = df_empresas['RIF'].isin(mapa_clientes.index)

    # Aplicar unificación solo donde coincide
    df_empresas.loc[coincide_rif, 'NOMCLIE'] = df_empresas.loc[coincide_rif, 'RIF'].map(mapa_clientes)

    # NUEVA COLUMNA: indicador de unificación
    df_empresas['UNIFICADO'] = coincide_rif.map({True: 'SI', False: 'NO'})

    coincidencias = coincide_rif.sum()
    print(f"✅ Unificación completada. Coincidencias por RIF encontradas: {coincidencias}")

    # --- AÑADIR PREFIJO '.' A CLIENTES ESPECÍFICOS ---
    print("\n--- Ajustando nombres seleccionados: añadiendo prefijo '.' a clientes específicos ---")
    clientes_a_prefijar = [
        'ACEROS DE VALENCIA, C.A',
        'ACEROS DE VALENCIA, C.A.',
        'EL PUNTO DEL HIERRO, C.A.',
        'EL PUNTO DEL HIERRO C.A',
        'FERREACEVAL, C.A.',
        'HIERRO EL ROBLE CCS, C.A.',
        'HIERRO METALES EL ROBLE, C.A.',
        'HIERROS PORTUGUESA, C.A.',
        'FERRETERIA PUNTO DEL HIERRO, C.A.',
        'FERRETERIA PUNTO DEL HIERRO, C.A',
        'CARLOS ANDRES ACEVEDO TOBO',
        'JUAN ALBERTO REINA OLAYA',  # FIX: coma agregada
        'CONSTRUCIONES BARINAS',
        'EL PUNTO DEL HIERRO BARINAS., C.A',
        'EL PUNTO DEL HIERRO BARINAS, C.A',
        'CONSTRUYE ALIANZA, C.A.',
        'CONSTRUYE ALIANZA C.A.',
        'CONSTRUYE ALIANZA, CA',
        'CONSTRUYE ALIANZA CA',  # FIX: coma agregada
        'CONSTRUYE ALIANZA., C.A.',
    ]

    def add_dot_if_matches(name):
        if pd.isna(name):
            return name
        name_str = str(name).strip()
        if name_str in clientes_a_prefijar:
            return name_str if name_str.startswith('.') else '.' + name_str
        return name_str

    df_empresas['NOMCLIE'] = df_empresas['NOMCLIE'].apply(add_dot_if_matches)
    print("✅ Prefijos aplicados (si correspondía).")

    # --- MODIFICACIÓN 2: Sustituir nombres de vendedores en NOMVEND ---
    print("\n--- Ajustando columna 'NOMVEND' ---")
    if 'NOMVEND' in df_empresas.columns:
        reemplazo_vendedores = {
            'CARLOS ACEVEDO': 'CARLOS ANDRES ACEVEDO TOBO',
            'JUAN REINA': 'JUAN ALBERTO REINA OLAYA'
        }

        df_empresas['NOMVEND'] = (
            df_empresas['NOMVEND']
            .astype(str)
            .str.strip()
            .str.replace(r'\s+', ' ', regex=True)
            .replace(reemplazo_vendedores)
        )
        print("✅ Nombres de vendedores sustituidos en 'NOMVEND'.")
    else:
        print("⚠️ Advertencia: No existe la columna 'NOMVEND' en el archivo.")

    print("\n--- Vista previa con columnas clave actualizadas ---")
    columnas_preview = [c for c in ['RIF', 'NOMCLIE_ORIGINAL', 'NOMCLIE', 'UNIFICADO', 'NOMVEND'] if c in df_empresas.columns]
    print(df_empresas[columnas_preview].head(10))

    # --- BLOQUE DE GUARDADO ---
    print("\n--- GUARDADO DE PROGRESO: Creando archivo Excel para el Paso 5 ---")
    nombre_archivo_salida = 'resultado_paso_5.xlsx'
    try:
        print(f"🔄 Guardando en '{nombre_archivo_salida}'...")
        df_empresas.to_excel(nombre_archivo_salida, index=False, sheet_name='Paso_5_Unificado')
        print(f"✅ ¡Éxito! El archivo '{nombre_archivo_salida}' ha sido creado.")
    except Exception as e:
        print(f"❌ Ocurrió un error al guardar el archivo: {e}")

    print("\n--- FIN DEL SCRIPT PARA PASO 5 ---")

def etapa_6():
    import pandas as pd
    import numpy as np
    import re

    print("--- INICIO: Ejecutando Paso 6 (CU con EMP abreviado y garantía de unicidad '-A') ---")

    # --- CARGA DEL PROGRESO ANTERIOR ---
    INPUT_FILE = 'resultado_paso_5.xlsx'
    OUTPUT_FILE = 'resultado_paso_6.xlsx'

    print(f"\n--- CARGANDO PROGRESO: Leyendo '{INPUT_FILE}' ---")
    try:
        columnas_como_texto = {
            'DOCUMENTO': str,
            'CODPROD': str,
            'CODVEND': str,
            'RIF': str,
            'FACTAFECT': str
        }
        df_empresas = pd.read_excel(INPUT_FILE, dtype=columnas_como_texto)

        # FECHA ahora puede venir con hora. Parseo robusto:
        if 'FECHA' in df_empresas.columns:
            df_empresas['FECHA'] = pd.to_datetime(
                df_empresas['FECHA'],
                errors='coerce',
                dayfirst=True
            )

        print(f"✅ Archivo '{INPUT_FILE}' cargado correctamente. Filas: {len(df_empresas)}")
    except FileNotFoundError:
        print(f"❌ Error: No se encontró el archivo '{INPUT_FILE}'. Asegúrate de ejecutar el paso anterior.")
        raise SystemExit(1)

    # --- NORMALIZACIONES BÁSICAS ---
    columnas_clave = ['TIPDOC', 'CAJA', 'DOCUMENTO']
    for col in columnas_clave:
        if col in df_empresas.columns:
            df_empresas[col] = df_empresas[col].fillna('').astype(str).str.strip()

    # Manejo de CAJA: intentar entero, sino dejar como string
    if 'CAJA' in df_empresas.columns:
        caja_num = pd.to_numeric(df_empresas['CAJA'], errors='coerce')
        df_empresas.loc[caja_num.notna(), 'CAJA'] = caja_num[caja_num.notna()].astype(int).astype(str)
        df_empresas.loc[caja_num.isna(), 'CAJA'] = df_empresas.loc[caja_num.isna(), 'CAJA'].fillna('').astype(str).str.strip()

    # Asegurar columnas monetarias numéricas
    for col in ['CONTADO', 'CREDITO', 'ANTICIPOCON', 'ANTICIPOCXC']:
        if col in df_empresas.columns:
            df_empresas[col] = pd.to_numeric(df_empresas[col], errors='coerce').fillna(0)
        else:
            df_empresas[col] = 0.0

    # ------------------ MAPEADO DE EMP A ABREVIATURA ------------------
    emp_map = {
        'ACEVAL': 'ACEV',
        'BARQUISIMETO': 'BQTO',
        'FERREACEVAL': 'FERRE',
        'CARACAS': 'CCS',
        'MARACAY': 'MCY',
        'PORTUGUESA': 'PORT',
        'BARCELONA': 'BLN',
        'CONSTALIANZA': 'CONSTALIANZA'
    }

    def get_emp_abbr(val) -> str:
        if pd.isna(val):
            return ''
        s = str(val).strip().upper()
        if s in emp_map:
            return emp_map[s]
        for k, v in emp_map.items():
            if k in s:
                return v
        return ''

    if 'EMP' in df_empresas.columns:
        df_empresas['EMP_ABBR'] = df_empresas['EMP'].apply(get_emp_abbr)
    else:
        df_empresas['EMP_ABBR'] = ''

    # ------------------ Obtener FISCAL/GUIA ------------------
    posibles_fiscal_cols = ['FISCAL/GUIA', 'FISCAL', 'GUIA']
    fiscal_col = next((c for c in posibles_fiscal_cols if c in df_empresas.columns), None)

    def extract_fiscal_letter(val) -> str:
        if pd.isna(val):
            return ''
        s = str(val).strip().upper()
        m = re.search(r'[A-Z]', s)
        return m.group(0) if m else ''

    if fiscal_col:
        df_empresas['FISCAL_LETRA'] = df_empresas[fiscal_col].apply(extract_fiscal_letter)
    else:
        df_empresas['FISCAL_LETRA'] = ''

    # ------------------ Formateo de FECHA para la CU (SIN HORA) ------------------
    if 'FECHA' in df_empresas.columns and pd.api.types.is_datetime64_any_dtype(df_empresas['FECHA']):
        df_empresas['FECHA_FOR_CU'] = df_empresas['FECHA'].dt.strftime('%d/%m/%Y').fillna('')
    else:
        df_empresas['FECHA_FOR_CU'] = ''

    # ------------------ Determinar código de condición CO/CR/DU/AN/ANCR/ANCO ------------------
    group_cols_cond = ['TIPDOC', 'EMP', 'FECHA_FOR_CU', 'CAJA', 'DOCUMENTO']
    if fiscal_col:
        group_cols_cond.insert(2, fiscal_col)  # TIPDOC, EMP, FISCAL/GUIA, FECHA, CAJA, DOCUMENTO

    cond_summary = (
        df_empresas
        .groupby(group_cols_cond)[['CONTADO', 'CREDITO', 'ANTICIPOCON', 'ANTICIPOCXC']]
        .sum(min_count=1)
        .reset_index()
    )

    def choose_cond(row):
        anticipo_con = row.get('ANTICIPOCON', 0) or 0
        anticipo_cxc = row.get('ANTICIPOCXC', 0) or 0
        contado_val = row.get('CONTADO', 0) or 0
        credito_val = row.get('CREDITO', 0) or 0

        anticipo = (anticipo_con + anticipo_cxc) > 0
        contado = contado_val > 0
        credito = credito_val > 0

        if anticipo and credito:
            return 'ANCR'
        if anticipo and contado:
            return 'ANCO'
        if anticipo:
            return 'AN'

        if contado and not credito:
            return 'CO'
        if credito and not contado:
            return 'CR'
        if credito and contado:
            return 'DU'

        return ''

    def build_cond_key(row):
        if fiscal_col:
            return (
                row.get('TIPDOC', ''),
                row.get('EMP', ''),
                row.get(fiscal_col, ''),
                row.get('FECHA_FOR_CU', ''),
                row.get('CAJA', ''),
                row.get('DOCUMENTO', '')
            )
        return (
            row.get('TIPDOC', ''),
            row.get('EMP', ''),
            row.get('FECHA_FOR_CU', ''),
            row.get('CAJA', ''),
            row.get('DOCUMENTO', '')
        )

    cond_map = {build_cond_key(r): choose_cond(r) for _, r in cond_summary.iterrows()}

    df_empresas['COND_COD'] = df_empresas.apply(
        lambda r: cond_map.get(build_cond_key(r), ''),
        axis=1
    )

    # ------------------ Bandera RET por RETENCION > 0 ------------------
    df_empresas['RET_FLAG'] = np.where(
        pd.to_numeric(df_empresas.get('RETENCION', 0), errors='coerce').fillna(0) > 0,
        'RET',
        ''
    )

    # ------------------ Construcción de la CU_BASE (sin sufijo) ------------------
    def build_fiscal_cond(ret_flag: str, fiscal: str, cond: str) -> list:
        fiscal_cond = (fiscal + cond).strip()
        parts = []
        if fiscal_cond:
            parts.append(fiscal_cond)
        if ret_flag:
            parts.append(ret_flag)
        return parts

    def build_cu_base(row):
        tipodoc = str(row.get('TIPDOC', '')).strip()
        emp_abbr = str(row.get('EMP_ABBR', '')).strip()
        fiscal = str(row.get('FISCAL_LETRA', '')).strip()
        cond = str(row.get('COND_COD', '')).strip()
        ret_flag = str(row.get('RET_FLAG', '')).strip()
        caja = str(row.get('CAJA', '')).strip()
        documento = str(row.get('DOCUMENTO', '')).strip()
        fecha = str(row.get('FECHA_FOR_CU', '')).strip()  # solo fecha, sin hora

        parts = [tipodoc]
        if emp_abbr:
            parts.append(emp_abbr)
        parts.extend(build_fiscal_cond(ret_flag, fiscal, cond))
        parts.extend([caja, documento, fecha])

        return '-'.join([p for p in parts if p != ''])

    df_empresas['CU_BASE'] = df_empresas.apply(build_cu_base, axis=1)

    # ------------------ Asignar sufijo '-A' a la PRIMERA fila por CU_BASE ------------------
    is_first_per_cu = ~df_empresas.duplicated(subset=['CU_BASE'], keep='first')
    df_empresas['CU'] = df_empresas['CU_BASE'].astype(str)
    df_empresas.loc[is_first_per_cu, 'CU'] = df_empresas.loc[is_first_per_cu, 'CU'] + '-A'
    df_empresas.loc[~is_first_per_cu, 'CU'] = (
        df_empresas.loc[~is_first_per_cu, 'CU']
        .astype(str)
        .str.replace(r'-A$', '', regex=True)
    )

    print("✅ CU_BASE construida y sufijos '-A' asignados (incluye EMP_ABBR, FECHA sin hora, condición y RET).")

    # ------------------ VALIDACIÓN ADICIONAL: garantizar unicidad '-A' por CU_BASE ------------------
    mask_A = df_empresas['CU'].str.endswith('-A', na=False)
    counts_A_per_base = df_empresas.loc[mask_A, 'CU_BASE'].value_counts()
    conflict_bases = counts_A_per_base[counts_A_per_base > 1]

    if not conflict_bases.empty:
        print("⚠️ Se detectaron CU_BASE con más de una fila marcada con '-A'. Se aplicará corrección conservadora.")
        for base, _ in conflict_bases.items():
            idxs = df_empresas[(df_empresas['CU_BASE'] == base) & df_empresas['CU'].str.endswith('-A')].index.tolist()
            for idx in idxs[1:]:
                df_empresas.at[idx, 'CU'] = re.sub(r'-A$', '', str(df_empresas.at[idx, 'CU']))
        print(f"✅ Corrección aplicada a {len(conflict_bases)} CU_BASE con conflicto.")
    else:
        print("✅ Validación: No se encontraron conflictos de '-A' por CU_BASE.")

    # ------------------ Reporte resumido ------------------
    total_cu_base = df_empresas['CU_BASE'].nunique()
    total_cu_with_A = df_empresas['CU'].str.endswith('-A', na=False).sum()
    print(f"\nResumen: CU_BASE únicas = {total_cu_base}, filas con '-A' = {total_cu_with_A}")

    # --- Corrección: FACTAFECT vacío debe permanecer vacío ---
    if 'FACTAFECT' in df_empresas.columns:
        df_empresas['FACTAFECT'] = df_empresas['FACTAFECT'].replace(
            ['nan', 'NaN', 'None', '<NA>', ''],
            pd.NA
        )

    # --- GUARDADO ---
    print(f"\n--- GUARDADO: escribiendo '{OUTPUT_FILE}' ---")
    try:
        df_save = df_empresas.copy()

        # Guardar FECHA con hora (si existe), pero CU ya usa FECHA_FOR_CU sin hora
        if 'FECHA' in df_save.columns and pd.api.types.is_datetime64_any_dtype(df_save['FECHA']):
            df_save['FECHA'] = df_save['FECHA'].dt.strftime('%d/%m/%Y %H:%M:%S')

        df_save.to_excel(OUTPUT_FILE, index=False, sheet_name='Paso_6_CU')
        print(f"✅ Archivo '{OUTPUT_FILE}' creado correctamente.")
    except Exception as e:
        print(f"❌ Error al guardar '{OUTPUT_FILE}': {e}")

    print("\n--- FIN DEL SCRIPT PARA PASO 6 ---")

def etapa_7():
    import pandas as pd
    import numpy as np

    print("--- INICIO: Ejecutando Paso 7 (Unificación DESCRIP/FAM desde DESCRIP_UNICAS) ---")

    # --- CARGA DEL PROGRESO ANTERIOR ---
    print("\n--- CARGANDO PROGRESO: Leyendo 'resultado_paso_6.xlsx' ---")
    archivo_anterior = 'resultado_paso_6.xlsx'
    try:
        columnas_como_texto = {
            'DOCUMENTO': str,
            'CODPROD': str,
            'CODVEND': str,
            'RIF': str,
            'FACTAFECT': str,
            'EMP': str,
            'TIPO': str
        }
        df_empresas = pd.read_excel(archivo_anterior, dtype=columnas_como_texto)

        print(f"✅ Archivo '{archivo_anterior}' cargado. Contiene {len(df_empresas)} filas.")

    except FileNotFoundError:
        print(f"❌ Error: No se encontró el archivo '{archivo_anterior}'.")
        raise SystemExit(1)

    # --- Limpieza básica de textos sin forzar nulos a 'nan' ---
    for col in df_empresas.select_dtypes(include=['object']).columns:
        df_empresas[col] = df_empresas[col].where(
            df_empresas[col].isna(),
            df_empresas[col].str.strip().str.replace(r'\s+', ' ', regex=True)
        )

    # --- PASO 7: Actualizar DESCRIP y FAM desde DESCRIP_UNICAS.xlsx ---
    print("\n--- PASO 7: Actualizando 'DESCRIP' y 'FAM' desde 'DESCRIP_UNICAS.xlsx' ---")
    try:
        df_descr_unicas = pd.read_excel(
            'DESCRIP_UNICAS.xlsx',
            dtype={'EMP': str, 'TIPO': str, 'CODPROD': str}
        )
        print(f"✅ Archivo 'DESCRIP_UNICAS.xlsx' cargado. Contiene {len(df_descr_unicas)} filas.")

        # Validar columnas obligatorias
        columnas_obligatorias = ['EMP', 'TIPO', 'CODPROD', 'DESCRIP', 'FAM']
        faltantes = [c for c in columnas_obligatorias if c not in df_descr_unicas.columns]
        if faltantes:
            print(f"❌ Error: Faltan columnas en DESCRIP_UNICAS.xlsx: {faltantes}")
            raise SystemExit(1)

        # Limpieza de espacios
        for col in df_descr_unicas.select_dtypes(include=['object']).columns:
            df_descr_unicas[col] = df_descr_unicas[col].where(
                df_descr_unicas[col].isna(),
                df_descr_unicas[col].str.strip().str.replace(r'\s+', ' ', regex=True)
            )

        # Claves de unión obligatorias (prelación)
        claves_union = ['EMP', 'TIPO', 'CODPROD']

        # Anti-explosión por duplicados en archivo maestro
        duplicados_mask = df_descr_unicas.duplicated(subset=claves_union, keep=False)
        if duplicados_mask.any():
            print("\n⚠️ ADVERTENCIA: Se encontraron duplicados por EMP+TIPO+CODPROD en DESCRIP_UNICAS.xlsx.")
            print("   Se conservará la PRIMERA aparición para evitar multiplicación de filas.")
            df_descr_unicas = df_descr_unicas.drop_duplicates(subset=claves_union, keep='first')
            print(f"✅ Corrección aplicada. Filas únicas para unión: {len(df_descr_unicas)}")

        # Renombrar columnas origen para merge
        df_descr_unicas = df_descr_unicas.rename(columns={
            'DESCRIP': 'DESCRIP_NUEVA',
            'FAM': 'FAM_NUEVA'
        })

        # Merge seguro
        print("\n🔄 Realizando unión segura por EMP+TIPO+CODPROD...")
        df_empresas = pd.merge(
            df_empresas,
            df_descr_unicas[['EMP', 'TIPO', 'CODPROD', 'DESCRIP_NUEVA', 'FAM_NUEVA']],
            on=claves_union,
            how='left'
        )
        print(f"✅ Unión completada. El DataFrame final mantiene {len(df_empresas)} filas.")

        # Columna de control
        hay_descr = df_empresas['DESCRIP_NUEVA'].notna() if 'DESCRIP_NUEVA' in df_empresas.columns else False
        hay_fam = df_empresas['FAM_NUEVA'].notna() if 'FAM_NUEVA' in df_empresas.columns else False
        df_empresas['CAMBIO_EJECUTADO'] = np.where(hay_descr | hay_fam, 'SI', 'NO')

        filas_a_actualizar = df_empresas['CAMBIO_EJECUTADO'] == 'SI'

        # Sustitución en columnas destino DESCRIP y FAM
        if 'DESCRIP' not in df_empresas.columns:
            df_empresas['DESCRIP'] = pd.NA
        if 'FAM' not in df_empresas.columns:
            df_empresas['FAM'] = pd.NA

        df_empresas.loc[filas_a_actualizar & df_empresas['DESCRIP_NUEVA'].notna(), 'DESCRIP'] = \
            df_empresas.loc[filas_a_actualizar & df_empresas['DESCRIP_NUEVA'].notna(), 'DESCRIP_NUEVA']

        df_empresas.loc[filas_a_actualizar & df_empresas['FAM_NUEVA'].notna(), 'FAM'] = \
            df_empresas.loc[filas_a_actualizar & df_empresas['FAM_NUEVA'].notna(), 'FAM_NUEVA']

        # Eliminar auxiliares
        df_empresas.drop(columns=['DESCRIP_NUEVA', 'FAM_NUEVA'], inplace=True)

        # Reordenar CAMBIO_EJECUTADO junto a FAM
        lista_columnas = df_empresas.columns.tolist()
        if 'CAMBIO_EJECUTADO' in lista_columnas and 'FAM' in lista_columnas:
            lista_columnas.remove('CAMBIO_EJECUTADO')
            pos_fam = lista_columnas.index('FAM')
            lista_columnas.insert(pos_fam + 1, 'CAMBIO_EJECUTADO')
            df_empresas = df_empresas[lista_columnas]

        total_si = (df_empresas['CAMBIO_EJECUTADO'] == 'SI').sum()
        total_no = (df_empresas['CAMBIO_EJECUTADO'] == 'NO').sum()
        print(f"✅ Actualización completada. CAMBIO_EJECUTADO -> SI: {total_si}, NO: {total_no}")

    except FileNotFoundError:
        print("❌ Error: No se encontró el archivo 'DESCRIP_UNICAS.xlsx'.")
        raise SystemExit(1)
    except Exception as e:
        print(f"❌ Ocurrió un error inesperado durante el Paso 7: {e}")
        raise SystemExit(1)

    # --- Mantener FACTAFECT vacío (no 'nan' texto) ---
    if 'FACTAFECT' in df_empresas.columns:
        df_empresas['FACTAFECT'] = df_empresas['FACTAFECT'].replace(
            ['nan', 'NaN', 'None', '<NA>', ''],
            pd.NA
        )

    # --- GUARDADO ---
    print("\n--- GUARDADO DE PROGRESO: Creando archivo Excel para el Paso 7 ---")
    nombre_archivo_salida = 'resultado_paso_7.xlsx'
    try:
        print(f"🔄 Guardando en '{nombre_archivo_salida}'...")
        df_empresas.to_excel(nombre_archivo_salida, index=False, sheet_name='Paso_7_Descrip_Fam')
        print(f"✅ ¡Éxito! El archivo '{nombre_archivo_salida}' ha sido creado.")

    except Exception as e:
        print(f"❌ Ocurrió un error al guardar el archivo: {e}")

    print("\n--- FIN DEL SCRIPT PARA PASO 7 ---")

def etapa_8_y_9():
    import pandas as pd
    import numpy as np

    print("--- INICIO: Ejecutando Pasos 8 y 9 (cruce por FECHA-DÍA) ---")

    # --- CARGA DEL PROGRESO ANTERIOR ---
    print("\n--- CARGANDO PROGRESO: Leyendo 'resultado_paso_7.xlsx' ---")
    archivo_anterior = 'resultado_paso_7.xlsx'
    try:
        columnas_como_texto = {
            'DOCUMENTO': str,
            'CODPROD': str,
            'CODVEND': str,
            'RIF': str,
            'FACTAFECT': str
        }
        df_empresas = pd.read_excel(archivo_anterior, dtype=columnas_como_texto)

        if 'FECHA' not in df_empresas.columns:
            print("❌ Error: No existe la columna 'FECHA' en resultado_paso_7.xlsx.")
            raise SystemExit(1)

        # Parseo robusto de FECHA (admite varios formatos)
        fecha_tmp = pd.to_datetime(
            df_empresas['FECHA'],
            format='%Y-%m-%d %H:%M:%S.%f',
            errors='coerce'
        )
        mask_na = fecha_tmp.isna()
        if mask_na.any():
            fecha_tmp.loc[mask_na] = pd.to_datetime(
                df_empresas.loc[mask_na, 'FECHA'],
                errors='coerce',
                dayfirst=True
            )

        df_empresas['FECHA'] = fecha_tmp
        df_empresas['FECHA_DIA'] = df_empresas['FECHA'].dt.normalize()

        print(f"✅ Archivo '{archivo_anterior}' cargado. Contiene {len(df_empresas)} filas.")

    except FileNotFoundError:
        print(f"❌ Error: No se encontró el archivo '{archivo_anterior}'.")
        raise SystemExit(1)

    # --- PASOS 8 y 9: Incorporar Tasas desde TASAS.xlsx ---
    print("\n--- PASOS 8 & 9: Añadiendo Tasa Paralela y Tasa BCV desde TASAS.xlsx ---")

    try:
        archivo_tasas = 'TASAS.xlsx'
        print(f"🔄 Cargando archivo '{archivo_tasas}'...")
        xls = pd.ExcelFile(archivo_tasas)
        print(f"   - Hojas encontradas: {xls.sheet_names}")

        if 'PAR' not in xls.sheet_names or 'BCV' not in xls.sheet_names:
            raise Exception("El archivo TASAS.xlsx debe tener las hojas 'PAR' y 'BCV'.")

        # --- Hoja PAR ---
        df_par = pd.read_excel(xls, sheet_name='PAR')
        col_tasa_par = 'TASA PARALELA'
        col_fecha_par = 'FECHA PARALELA'

        for c in [col_tasa_par, col_fecha_par]:
            if c not in df_par.columns:
                raise KeyError(f"Falta la columna '{c}' en hoja PAR.")

        # Parseo de fecha PAR con formato esperado + fallback
        fecha_par = pd.to_datetime(df_par[col_fecha_par], format='%Y-%m-%d %H:%M:%S.%f', errors='coerce')
        mask_par_na = fecha_par.isna()
        if mask_par_na.any():
            fecha_par.loc[mask_par_na] = pd.to_datetime(df_par.loc[mask_par_na, col_fecha_par], errors='coerce', dayfirst=True)

        df_par[col_fecha_par] = fecha_par
        df_par[col_tasa_par] = pd.to_numeric(df_par[col_tasa_par], errors='coerce')
        df_par['FECHA_DIA'] = df_par[col_fecha_par].dt.normalize()

        mapa_par = (
            df_par
            .dropna(subset=['FECHA_DIA', col_tasa_par])
            .drop_duplicates(subset=['FECHA_DIA'], keep='last')
            [['FECHA_DIA', col_tasa_par]]
            .rename(columns={col_tasa_par: 'TasaEspAprox'})
        )

        # --- Hoja BCV ---
        df_bcv = pd.read_excel(xls, sheet_name='BCV')
        col_tasa_bcv = 'TASA BCV'
        col_fecha_bcv = 'FECHA BCV'

        for c in [col_tasa_bcv, col_fecha_bcv]:
            if c not in df_bcv.columns:
                raise KeyError(f"Falta la columna '{c}' en hoja BCV.")

        # Parseo de fecha BCV con formato esperado + fallback
        fecha_bcv = pd.to_datetime(df_bcv[col_fecha_bcv], format='%Y-%m-%d %H:%M:%S.%f', errors='coerce')
        mask_bcv_na = fecha_bcv.isna()
        if mask_bcv_na.any():
            fecha_bcv.loc[mask_bcv_na] = pd.to_datetime(df_bcv.loc[mask_bcv_na, col_fecha_bcv], errors='coerce', dayfirst=True)

        df_bcv[col_fecha_bcv] = fecha_bcv
        df_bcv[col_tasa_bcv] = pd.to_numeric(df_bcv[col_tasa_bcv], errors='coerce')
        df_bcv['FECHA_DIA'] = df_bcv[col_fecha_bcv].dt.normalize()

        mapa_bcv = (
            df_bcv
            .dropna(subset=['FECHA_DIA', col_tasa_bcv])
            .drop_duplicates(subset=['FECHA_DIA'], keep='last')
            [['FECHA_DIA', col_tasa_bcv]]
            .rename(columns={col_tasa_bcv: 'TasaBCV'})
        )

        # --- Merge por FECHA_DIA ---
        df_empresas = df_empresas.merge(mapa_par, on='FECHA_DIA', how='left')
        df_empresas = df_empresas.merge(mapa_bcv, on='FECHA_DIA', how='left')

        print("✅ Columnas 'TasaEspAprox' y 'TasaBCV' añadidas por coincidencia de fecha (día).")

    except Exception as e:
        print(f"\n⚠️ Advertencia: No se pudieron añadir las tasas de cambio. Error: {e}")
        print("   Se crearán las columnas 'TasaEspAprox' y 'TasaBCV' vacías para continuar.")
        if 'TasaEspAprox' not in df_empresas.columns:
            df_empresas['TasaEspAprox'] = np.nan
        if 'TasaBCV' not in df_empresas.columns:
            df_empresas['TasaBCV'] = np.nan

    # --- Verificación ---
    print("\n--- Verificación de las nuevas columnas de tasas: ---")
    col_tasa_original = 'TASAS' if 'TASAS' in df_empresas.columns else ('TASA' if 'TASA' in df_empresas.columns else None)
    columnas_a_ver = ['FECHA', 'TasaEspAprox', 'TasaBCV']
    if col_tasa_original:
        columnas_a_ver.insert(1, col_tasa_original)

    print(df_empresas[columnas_a_ver].head(10))

    print("\n--- Resumen de valores nulos en tasas ---")
    print(f"TasaEspAprox NaN: {df_empresas['TasaEspAprox'].isna().sum()}")
    print(f"TasaBCV NaN: {df_empresas['TasaBCV'].isna().sum()}")

    # Fechas sin match (auditoría rápida)
    falt_par = df_empresas[df_empresas['TasaEspAprox'].isna()]['FECHA_DIA'].dropna().drop_duplicates()
    falt_bcv = df_empresas[df_empresas['TasaBCV'].isna()]['FECHA_DIA'].dropna().drop_duplicates()

    if len(falt_par) > 0 or len(falt_bcv) > 0:
        print("\n⚠️ Fechas sin tasa encontrada (muestra hasta 10):")
        if len(falt_par) > 0:
            print("  - PAR faltantes:", list(falt_par.head(10).dt.strftime('%d/%m/%Y')))
        if len(falt_bcv) > 0:
            print("  - BCV faltantes:", list(falt_bcv.head(10).dt.strftime('%d/%m/%Y')))

    # --- BLOQUE DE GUARDADO ---
    print("\n--- GUARDADO DE PROGRESO: Creando archivo Excel para los Pasos 8 y 9 ---")
    nombre_archivo_salida = 'resultado_pasos_8_y_9.xlsx'
    try:
        df_para_guardar = df_empresas.copy()

        # Mantener FECHA con hora si existe
        if 'FECHA' in df_para_guardar.columns and pd.api.types.is_datetime64_any_dtype(df_para_guardar['FECHA']):
            df_para_guardar['FECHA'] = df_para_guardar['FECHA'].dt.strftime('%d/%m/%Y %H:%M:%S')

        # FECHA_DIA como apoyo visual (opcional). Si no la quieres en salida, la quitamos.
        if 'FECHA_DIA' in df_para_guardar.columns and pd.api.types.is_datetime64_any_dtype(df_para_guardar['FECHA_DIA']):
            df_para_guardar['FECHA_DIA'] = df_para_guardar['FECHA_DIA'].dt.strftime('%d/%m/%Y')

        print(f"🔄 Guardando en '{nombre_archivo_salida}'...")
        df_para_guardar.to_excel(nombre_archivo_salida, index=False, sheet_name='Pasos_8_y_9_Tasas')
        print(f"✅ ¡Éxito! El archivo '{nombre_archivo_salida}' ha sido creado.")

    except Exception as e:
        print(f"❌ Ocurrió un error al guardar el archivo: {e}")

    print("\n--- FIN DEL SCRIPT PARA PASOS 8 Y 9 ---")

def etapa_10_a_13():
    import pandas as pd
    import numpy as np

    print("--- INICIO: Ejecutando la Consolidación Final (Versión Robusta) ---")

    # --- CARGA DEL PROGRESO ANTERIOR ---
    print("\n--- CARGANDO PROGRESO: Leyendo 'resultado_pasos_8_y_9.xlsx' ---")
    archivo_anterior = 'resultado_pasos_8_y_9.xlsx'
    try:
        columnas_como_texto = {
            'DOCUMENTO': str,
            'CODPROD': str,
            'CODVEND': str,
            'RIF': str,
            'CU': str,
            'FACTAFECT': str
        }
        df_empresas = pd.read_excel(archivo_anterior, dtype=columnas_como_texto)

        print(f"✅ Archivo '{archivo_anterior}' cargado. Contiene {len(df_empresas)} filas.")
    except FileNotFoundError:
        print(f"❌ Error: No se encontró el archivo '{archivo_anterior}'.")
        raise SystemExit(1)

    # --- CONSOLIDACIÓN FINAL (LÓGICA MEJORADA) ---
    print("\n--- CONSOLIDACIÓN FINAL: Aplicando criterio a Pagos y Tasas ---")
    try:
        # 1) Columnas a consolidar por MAX (como ya lo venías haciendo)
        columnas_pago = [
            'FACTAFECT', 'CAJAFECTADA', 'SALDOPENDIENT', 'DOLAREFECTCONT', 'DOLAEXTCONT',
            'BSEFECTCONT', 'OTROSBSCONT', 'ANTICIPOCON', 'DOLAREFECTCRED',
            'DOLAREXTCRED', 'BSEFECTCRED', 'OTROSBSCRED', 'ANTICIPOCXC',
            'CONTADO', 'CREDITO', 'IVA', 'RETENCION', 'DESCUENTO'
        ]

        columnas_a_consolidar = [col for col in columnas_pago if col in df_empresas.columns]
        print(f"   - Columnas que serán consolidadas (MAX): {columnas_a_consolidar}")

        # Convertir a numérico
        for col in columnas_a_consolidar:
            df_empresas[col] = pd.to_numeric(df_empresas[col], errors='coerce').fillna(0)

        # 2) TOTAL para sumar por CU_BASE
        if 'TOTAL' in df_empresas.columns:
            df_empresas['TOTAL'] = pd.to_numeric(df_empresas['TOTAL'], errors='coerce').fillna(0)
        else:
            print("⚠️ Advertencia: No existe la columna 'TOTAL'. Se creará en 0.")
            df_empresas['TOTAL'] = 0.0

        # 3) Crear CU_BASE
        print("\n🔄 Creando clave de agrupación 'CU_BASE'...")
        df_empresas['CU_BASE'] = df_empresas['CU'].astype(str).str.replace('-A$', '', regex=True)

        # 4) Resúmenes por CU_BASE
        print("🔄 Creando tabla resumen...")
        resumen_max = df_empresas.groupby('CU_BASE')[columnas_a_consolidar].agg('max')
        resumen_total_fact = (
            df_empresas.groupby('CU_BASE', as_index=True)['TOTAL']
            .sum(min_count=1)
            .to_frame(name='TOTAL FACT')
        )
        print("✅ Tabla resumen creada.")

        # 5) Limpiar columnas consolidadas en tabla principal
        print("🔄 Limpiando columnas en tabla principal...")
        df_empresas[columnas_a_consolidar] = 0.0

        # Crear/limpiar TOTAL FACT (solo se llenará en filas -A)
        df_empresas['TOTAL FACT'] = 0.0
        print("✅ Limpieza completada.")

        # 6) Escribir valores consolidados en filas -A
        print("🔄 Escribiendo valores consolidados en filas '-A'...")
        es_fila_principal = df_empresas['CU'].astype(str).str.endswith('-A', na=False)

        # MAX columns
        df_max_A = pd.merge(
            df_empresas.loc[es_fila_principal, ['CU_BASE']],
            resumen_max,
            on='CU_BASE',
            how='left'
        )
        df_max_A[columnas_a_consolidar] = df_max_A[columnas_a_consolidar].fillna(0)
        df_empresas.loc[es_fila_principal, columnas_a_consolidar] = df_max_A[columnas_a_consolidar].values

        # TOTAL FACT (SUM de TOTAL por CU_BASE)
        df_total_A = pd.merge(
            df_empresas.loc[es_fila_principal, ['CU_BASE']],
            resumen_total_fact,
            on='CU_BASE',
            how='left'
        )
        df_total_A['TOTAL FACT'] = pd.to_numeric(df_total_A['TOTAL FACT'], errors='coerce').fillna(0)
        df_empresas.loc[es_fila_principal, 'TOTAL FACT'] = df_total_A['TOTAL FACT'].values

        # Eliminar temporal
        df_empresas.drop(columns=['CU_BASE'], inplace=True)

        print("✅ ¡Consolidación completada! 'TOTAL FACT' calculado y colocado en filas '-A'.")

    except Exception as e:
        print(f"❌ Ocurrió un error inesperado durante la consolidación: {e}")
        raise SystemExit(1)

    # --- BLOQUE DE GUARDADO FINAL ---
    print("\n--- GUARDADO FINAL: Creando el archivo consolidado definitivo ---")
    nombre_archivo_salida = 'Analisis_Consolidado_Final.xlsx'
    try:
        df_para_guardar = df_empresas.copy()

        if 'FECHA' in df_para_guardar.columns and pd.api.types.is_datetime64_any_dtype(df_para_guardar['FECHA']):
            df_para_guardar['FECHA'] = df_para_guardar['FECHA'].dt.strftime('%d/%m/%Y')

        print(f"🔄 Guardando en '{nombre_archivo_salida}'...")
        df_para_guardar.to_excel(nombre_archivo_salida, index=False, sheet_name='Analisis_Final_Robusto')

        print("\n\n" + "=" * 60)
        print("🎉🎉🎉 ¡¡TRABAJO COMPLETADO, Aryana-ia!! 🎉🎉🎉")
        print(f"El archivo '{nombre_archivo_salida}' ha sido creado con éxito.")
        print("Contiene todos los datos limpios, procesados y consolidados.")
        print("=" * 60 + "\n")

    except Exception as e:
        print(f"❌ Ocurrió un error al guardar el archivo final: {e}")

    print("--- FIN DEL SCRIPT ---")

def etapa_tipo_venta_cashea():
    import pandas as pd
    import numpy as np

    print("--- INICIO: Paso TIPO DE VENTA (SOCIOS / NETAS / CASHEA) ---")

    archivo_entrada = 'Analisis_Consolidado_Final.xlsx'
    archivo_salida = 'Analisis_Consolidado_Final_con_tipo_venta.xlsx'

    try:
        df = pd.read_excel(archivo_entrada, dtype={
            'NOMVEND': str, 'NOMCLIE': str, 'DOCUMENTO': str,
            'CODPROD': str, 'CODVEND': str, 'RIF': str,
            'CU': str, 'FACTAFECT': str
        })
        print(f"✅ Archivo '{archivo_entrada}' cargado. Filas: {len(df)}")
    except FileNotFoundError:
        print(f"❌ No se encontró el archivo '{archivo_entrada}'.")
        raise SystemExit(1)

    # Limpieza de espacios
    for col in ['NOMVEND', 'NOMCLIE']:
        if col not in df.columns:
            print(f"❌ Falta columna requerida: {col}")
            raise SystemExit(1)
        df[col] = df[col].where(
            df[col].isna(),
            df[col].astype(str).str.strip().str.replace(r'\s+', ' ', regex=True)
        )

    # ── Reglas SOCIOS ──
    vendedores_socios = {
        'CARLOS ANDRES ACEVEDO TOBO',
        'JUAN ALBERTO REINA OLAYA'
    }

    # ── Reglas NETAS (exclusiones) ──
    clientes_excluir_netas = {
        '.ACEROS DE VALENCIA, C.A.',
        '.CARLOS ANDRES ACEVEDO TOBO',
        '.EL PUNTO DEL HIERRO C.A',
        '.FERREACEVAL, C.A.',
        '.FERRETERIA PUNTO DEL HIERRO, C.A',
        '.HIERRO EL ROBLE CCS, C.A.',
        '.HIERRO METALES EL ROBLE, C.A.',
        '.HIERROS PORTUGUESA, C.A.',
        '.JUAN ALBERTO REINA OLAYA',
        '.CONSTRUYE ALIANZA, C.A.'
    }

    vendedores_excluir_netas = {
        'ACEVAL',
        'CARLOS ANDRES ACEVEDO TOBO',
        'FACTURAS AL PERSONAL',
        'JUAN ALBERTO REINA OLAYA',
        'OD-DONACION'
    }

    # ── Columna nueva ──
    df['TIPO DE VENTA'] = 'FULL'

    # 1️⃣ SOCIOS
    mask_socios = df['NOMVEND'].isin(vendedores_socios)
    df.loc[mask_socios, 'TIPO DE VENTA'] = 'SOCIOS'

    # 2️⃣ NETAS (sin pisar SOCIOS)
    mask_netas = (
        ~df['NOMCLIE'].isin(clientes_excluir_netas)
        & ~df['NOMVEND'].isin(vendedores_excluir_netas)
    )
    df.loc[~mask_socios & mask_netas, 'TIPO DE VENTA'] = 'NETAS'

    # 3️⃣ CASHEA (prioridad máxima: pisa todo lo anterior)
    if 'CASHEA' not in df.columns:
        print("⚠️ Columna 'CASHEA' no encontrada. Se omite esta regla.")
    else:
        df['CASHEA'] = df['CASHEA'].astype(str).str.strip().str.upper()
        mask_cashea = df['CASHEA'] == 'SI'
        df.loc[mask_cashea, 'TIPO DE VENTA'] = 'CASHEA'
        print(f"📌 Filas marcadas como CASHEA: {mask_cashea.sum()}")

    print("\n--- Resumen ---")
    print(df['TIPO DE VENTA'].value_counts(dropna=False))

    # Guardar
    df.to_excel(archivo_salida, index=False, sheet_name='Analisis_Final')
    print(f"\n✅ Archivo guardado: '{archivo_salida}'")
    print("--- FIN ---")

def etapa_transformacion_final():
    import pandas as pd
    import sys
    from pathlib import Path

    # ── CONFIGURACIÓN ──
    ARCHIVO_ENTRADA = "Analisis_Consolidado_Final_con_tipo_venta.xlsx"
    ARCHIVO_SALIDA  = "Analisis_Final_DEF.xlsx"

    # Columnas que deben leerse como texto
    COLUMNAS_TEXTO = {
        "NOMVEND": str,
        "NOMCLIE": str,
        "DOCUMENTO": str,
        "CODPROD": str,
        "CODVEND": str,
        "RIF": str,
        "CU": str,
        "FACTAFECT": str,
    }

    # 1️⃣ Lectura del archivo


    ruta = Path(ARCHIVO_ENTRADA)

    if not ruta.exists():
        print(f"❌ No se encontró el archivo: {ARCHIVO_ENTRADA}")
        print("   Asegurate de que el archivo esté en la misma carpeta que este notebook.")
        sys.exit(1)

    print(f"📂 Leyendo: {ARCHIVO_ENTRADA} ...")
    df = pd.read_excel(ruta, engine="openpyxl", dtype=COLUMNAS_TEXTO)
    print(f"   ✔ {df.shape[0]:,} filas × {df.shape[1]} columnas cargadas.")
    print(f"\nColumnas forzadas como texto: {list(COLUMNAS_TEXTO.keys())}")
    print(f"\nColumnas originales:\n{df.columns.tolist()}")


    #2️⃣ Crear columna `COSTO X CANT`
    #Se calcula como **COSTO × CANT** y se inserta justo después de la columna `COSTO`.

    # Crear la columna
    df["COSTO X CANT"] = df["COSTO"] * df["CANT"]

    # Insertar justo después de COSTO
    idx_costo = df.columns.get_loc("COSTO")
    cols = df.columns.tolist()
    cols.remove("COSTO X CANT")
    cols.insert(idx_costo + 1, "COSTO X CANT")
    df = df[cols]

    print("✅ Columna 'COSTO X CANT' creada e insertada después de 'COSTO'.")
    print(f"\nMuestra:")
    print(df[["CANT", "COSTO", "COSTO X CANT"]].head(5).to_string(index=False))


    # 3️⃣ Negar columnas cuando `TIPDOC == "DEV"`
    #Cuando la columna `TIPDOC` sea **DEV**, las siguientes columnas deben ser **negativas**:
    #`CANT`, `PVPDOL`, `COSTO`, `COSTO X CANT`, `TOTAL`, `TOTAL FACT`, `PESO`, `TOTALPESO`, `PRECIOXKILO`, `TasaEspAprox`, `TasaBCV`

    # Columnas que deben ser negativas en filas DEV
    COLUMNAS_A_NEGAR = [
        "CANT",
        "PVPDOL",
        "COSTO",
        "COSTO X CANT",
        "TOTAL",
        "TOTAL FACT",
        "PESO",
        "TOTALPESO",
        "PRECIOXKILO",
        "TasaEspAprox",
        "TasaBCV",
    ]

    # Máscara para filas DEV (tolerante a espacios y mayúsculas/minúsculas)
    mask_dev = df["TIPDOC"].astype(str).str.strip().str.upper() == "DEV"
    n_dev = mask_dev.sum()

    for col in COLUMNAS_A_NEGAR:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")     # asegurar numérico
            df.loc[mask_dev, col] = -df.loc[mask_dev, col].abs()   # forzar negativo
        else:
            print(f"   ⚠️  Columna '{col}' no encontrada, se omite.")

    print(f"✅ {n_dev:,} filas DEV → columnas convertidas a negativo.")


    # 4️⃣ Reordenar y seleccionar columnas
    # Se conservan únicamente las columnas indicadas, en el orden final especificado.

    ORDEN_FINAL = [
        "TIPDOC",
        "EMP",
        "FISCAL/GUIA",
        "CODCLIE",
        "NOMCLIE",
        "RIF",
        "DOCUMENTO",
        "FECHA",
        "TIPO",
        "UND",
        "CODPROD",
        "DESCRIP",
        "FAM",
        "TASA",
        "CANT",
        "PVPDOL",
        "COSTO",
        "COSTO X CANT",
        "TIPOPVP",
        "DESCUENTO",
        "TOTAL",
        "TOTAL FACT",
        "CONTADO",
        "CREDITO",
        "IVA",
        "RETENCION",
        "PESO",
        "TOTALPESO",
        "PRECIOXKILO",
        "CAJA",
        "CODVEND",
        "NOMVEND",
        "FACTAFECT",
        "CAJAFECTADA",
        "SALDOPENDIENT",
        "DOLAREFECTCONT",
        "DOLAEXTCONT",
        "BSEFECTCONT",
        "OTROSBSCONT",
        "ANTICIPOCON",
        "DOLAREFECTCRED",
        "DOLAREXTCRED",
        "BSEFECTCRED",
        "OTROSBSCRED",
        "ANTICIPOCXC",
        "MARCA",
        "TONINICIAL",
        "TONFINAL",
        "CU",
        "TasaEspAprox",
        "TasaBCV",
        "TIPO DE VENTA",
    ]

    # Verificar columnas faltantes
    faltantes = [c for c in ORDEN_FINAL if c not in df.columns]
    if faltantes:
        print(f"⚠️  Columnas no encontradas en el archivo (se omitirán): {faltantes}")

    # Seleccionar y reordenar solo las que existen
    columnas_validas = [c for c in ORDEN_FINAL if c in df.columns]
    df = df[columnas_validas]

    # Columnas descartadas del original
    descartadas = sorted(set(cols) - set(ORDEN_FINAL))
    print(f"✅ Columnas reordenadas: {len(columnas_validas)} columnas seleccionadas.")
    if descartadas:
        print(f"   Columnas descartadas del original: {descartadas}")
    
    
    
    # Guardar resultado

    df.to_excel(ARCHIVO_SALIDA, index=False, sheet_name="Analisis_Final", engine="openpyxl")

    print(f"💾 Archivo guardado: {ARCHIVO_SALIDA}")
    print(f"   → {df.shape[0]:,} filas × {df.shape[1]} columnas")



    #Verificacion

    print(f"── COLUMNAS FINALES ({len(df.columns)}) ──")
    for i, col in enumerate(df.columns, 1):
        print(f"   {i:>2}. {col}")

    if n_dev > 0:
        muestra_cols = ["TIPDOC"] + [c for c in COLUMNAS_A_NEGAR if c in df.columns]
        print(f"\n── MUESTRA DE FILAS DEV (deben ser negativas) ──")
        print(df.loc[mask_dev, muestra_cols].head(5).to_string(index=False))

    print("\n🎉 ¡Proceso completado exitosamente!")

    

