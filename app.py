import streamlit as st
import pandas as pd
from docx import Document
import io
from datetime import datetime
import gspread

# --- CONFIGURACIÓN Y CONEXIÓN SEGURA A GOOGLE SHEETS ---
try:
    # Autenticación usando los "Secrets" de Streamlit
    gc = gspread.service_account_from_dict(st.secrets["gcp_service_account"])

    # Abrir las hojas de cálculo por su nombre
    sh_personas = gc.open("Base de Datos - Personas").sheet1
    sh_cursos = gc.open("Base de Datos - Cursos").sheet1

    # Convertir los datos a DataFrames de Pandas
    df_personas_full = pd.DataFrame(sh_personas.get_all_records())
    df_cursos = pd.DataFrame(sh_cursos.get_all_records())

    # Asegurarse de que DNI sea texto para las comparaciones
    df_personas_full['DNI'] = df_personas_full['DNI'].astype(str)

except Exception as e:
    st.error(f"Error al conectar con Google Sheets. Verifica la configuración de API y los Secrets: {e}")
    st.stop() # Detiene la ejecución si no se puede conectar

# --- FUNCIÓN PRINCIPAL QUE GENERA EL DOCUMENTO ---
def generar_documento(curso_elegido_df, dnis_a_procesar):
    cursantes_df = df_personas_full[df_personas_full['DNI'].isin(dnis_a_procesar)]
    st.info(f"Se encontraron {len(cursantes_df)} de {len(dnis_a_procesar)} DNI en la base de datos.")

    documento = Document('plantilla.docx')
    
    # Reemplazar marcadores del curso y otros datos
    datos_completos = curso_elegido_df.to_dict()
    todos_los_parrafos = list(documento.paragraphs)
    for tabla in documento.tables:
        for row in tabla.rows:
            for cell in row.cells:
                for parrafo in cell.paragraphs:
                    todos_los_parrafos.append(parrafo)

    for parrafo in todos_los_parrafos:
        for key, value in datos_completos.items():
            marcador = f"{{{{{key}}}}}"
            if marcador in parrafo.text:
                dato_a_reemplazar = ''
                if key in ['Fecha_Inicio', 'Fecha_Fin'] and pd.notna(value) and value != '':
                    try:
                        dato_a_reemplazar = pd.to_datetime(value).strftime('%d/%m/%Y')
                    except (ValueError, TypeError):
                        dato_a_reemplazar = str(value) # Si no es una fecha válida, lo dejamos como está
                else:
                    dato_a_reemplazar = str(value)
                parrafo.text = parrafo.text.replace(marcador, dato_a_reemplazar)
    
    # Encontrar el marcador de la tabla e insertarla
    parrafo_marcador = None
    for p in documento.paragraphs:
        if '{{TABLA_PARTICIPANTES}}' in p.text:
            parrafo_marcador = p
            break

    if parrafo_marcador:
        parrafo_marcador.text = ""
        tabla = documento.add_table(rows=1, cols=3)
        try: tabla.style = 'Tabla con cuadrícula'
        except KeyError:
            try: tabla.style = 'Table Grid'
            except KeyError: st.warning("Estilo de tabla no encontrado.")
        
        hdr_cells = tabla.rows[0].cells
        hdr_cells[0].text = 'Jerarquía'
        hdr_cells[1].text = 'DNI'
        hdr_cells[2].text = 'Nombre y Apellido'
        
        for index, persona in cursantes_df.iterrows():
            row_cells = tabla.add_row().cells
            row_cells[0].text = str(persona.get('Jerarquia', ''))
            row_cells[1].text = str(persona.get('DNI', ''))
            row_cells[2].text = str(persona.get('Nombre_Apellido', ''))
        
        parrafo_marcador._p.addnext(tabla._element)
    else:
        st.warning("ADVERTENCIA: No se encontró el marcador {{TABLA_PARTICIPANTES}} en la plantilla.")

    buffer = io.BytesIO()
    documento.save(buffer)
    buffer.seek(0)
    return buffer

# --- INTERFAZ DE LA PÁGINA WEB ---
st.title("Generador de Informes de Cursos 🚀")

lista_cursos = df_cursos['Nombre_Curso'].tolist()
curso_seleccionado_nombre = st.selectbox("Paso 1: Seleccione el curso", lista_cursos)
archivo_dni_subido = st.file_uploader("Paso 2: Suba el archivo `lista_dni.txt`", type="txt")

if st.button("Generar Documento"):
    if archivo_dni_subido is not None:
        with st.spinner('Procesando...'):
            dnis = archivo_dni_subido.getvalue().decode("utf-8").splitlines()
            dnis_limpios = [dni.strip() for dni in dnis if dni.strip()]

            curso_elegido_df = df_cursos[df_cursos['Nombre_Curso'] == curso_seleccionado_nombre].iloc[0]
            buffer_documento = generar_documento(curso_elegido_df, dnis_limpios)

            nombre_curso_corto = curso_seleccionado_nombre.replace(" ", "_")[:20]
            nombre_archivo = f"Informe_{nombre_curso_corto}_{datetime.now().strftime('%Y-%m-%d')}.docx"
            
            st.success("¡Documento generado con éxito!")
            st.download_button(
                label="Descargar Documento Word",
                data=buffer_documento,
                file_name=nombre_archivo,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
    else:
        st.error("Por favor, suba el archivo `lista_dni.txt`.")