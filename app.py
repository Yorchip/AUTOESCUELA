import sys
import subprocess

# Instalar automáticamente librerías requeridas si no están presentes
for pkg in ["openpyxl", "PyGithub"]:
    try:
        __import__(pkg if pkg != "PyGithub" else "github")
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", pkg])

import streamlit as st
import pandas as pd
import random
from pathlib import Path
import os
import io

# Configuración de página
st.set_page_config(
    page_title="Sistema de Exámenes Profesional",
    page_icon="📚",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Estilo CSS para ajustar imágenes de forma responsive (PC y Móvil)
st.markdown("""
    <style>
    div[data-testid="stImage"] img {
        max-height: 300px; /* Limita la altura en monitor para ver pregunta y respuestas sin scroll */
        width: auto; /* Ancho proporcional */
        max-width: 100%; /* Adapta el ancho al móvil sin salirse de la pantalla */
        object-fit: contain; /* Evita que la imagen se recorte o deforme */
        margin: 0 auto; /* Centra la imagen */
        display: block;
    }
    </style>
""", unsafe_allow_html=True)

ESTADISTICAS_FILE = 'estadisticas.xlsx'
DATOS_DIR = Path('datos')
IMAGENES_DIR = Path('imagenes')

# --- FUNCIONES DE CARGA Y ESTADÍSTICAS ---

def load_statistics():
    if Path(ESTADISTICAS_FILE).exists():
        try:
            df = pd.read_excel(ESTADISTICAS_FILE)
            estadisticas = {}
            for _, fila in df.iterrows():
                clave = f"{fila['tema']}:{fila['pregunta_corta']}"
                estadisticas[clave] = {
                    'tema': fila['tema'],
                    'pregunta': fila['pregunta_completa'],
                    'intentos': int(fila['intentos']),
                    'aciertos': int(fila['aciertos']),
                    'fallos': int(fila['fallos'])
                }
            return estadisticas
        except Exception as e:
            st.error(f"Error al cargar estadísticas: {e}")
            return {}
    return {}

def save_statistics(estadisticas):
    try:
        datos = []
        for clave, stats in estadisticas.items():
            pregunta_corta = str(stats['pregunta'])[:50].strip()
            datos.append({
                'tema': stats['tema'],
                'pregunta_corta': pregunta_corta,
                'pregunta_completa': stats['pregunta'],
                'intentos': stats['intentos'],
                'aciertos': stats['aciertos'],
                'fallos': stats['fallos'],
                'porcentaje_fallos': round((stats['fallos'] / stats['intentos']) * 100, 1) if stats['intentos'] > 0 else 0,
                'porcentaje_aciertos': round((stats['aciertos'] / stats['intentos']) * 100, 1) if stats['intentos'] > 0 else 0
            })
        
        df = pd.DataFrame(datos)
        if not df.empty:
            df = df.sort_values(['porcentaje_fallos', 'fallos'], ascending=[False, False])
        
        # 1. Guardar localmente en el servidor
        with pd.ExcelWriter(ESTADISTICAS_FILE, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Estadisticas', index=False)
        
        # 2. Guardar en GitHub de forma permanente
        if "GITHUB_TOKEN" in st.secrets:
            from github import Github, GithubException
            token = st.secrets["GITHUB_TOKEN"]
            g = Github(token)
            repo = g.get_repo("Yorchip/AUTOESCUELA")
            
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name='Estadisticas', index=False)
            
            content_bytes = buffer.getvalue()
            file_path = "estadisticas.xlsx"
            
            try:
                # Obtener el archivo existente y su SHA actual para actualizarlo
                contents = repo.get_contents(file_path)
                repo.update_file(
                    path=file_path,
                    message="📊 Actualización automática de estadísticas",
                    content=content_bytes,
                    sha=contents.sha
                )
            except GithubException as ge:
                if ge.status == 404:
                    # Crear el archivo solo si realmente no existe en GitHub
                    repo.create_file(
                        path=file_path,
                        message="📊 Creación automática de estadísticas",
                        content=content_bytes
                    )
                else:
                    st.error(f"Error al actualizar estadísticas en GitHub: {ge}")
    except Exception as e:
        st.error(f"Error al guardar estadísticas: {e}")

def update_statistics(pregunta_data, correcto):
    estadisticas = load_statistics()
    pregunta_corta = str(pregunta_data['pregunta'])[:50].strip()
    clave = f"{pregunta_data['tema']}:{pregunta_corta}"
    
    if clave not in estadisticas:
        estadisticas[clave] = {
            'tema': pregunta_data['tema'],
            'pregunta': pregunta_data['pregunta'],
            'intentos': 0,
            'aciertos': 0,
            'fallos': 0
        }
    
    estadisticas[clave]['intentos'] += 1
    if correcto:
        estadisticas[clave]['aciertos'] += 1
    else:
        estadisticas[clave]['fallos'] += 1
    
    save_statistics(estadisticas)

def get_available_topics():
    if not DATOS_DIR.exists():
        return []
    archivos = list(DATOS_DIR.glob('*.xlsx')) + list(DATOS_DIR.glob('*.csv')) + list(DATOS_DIR.glob('*.XLSX')) + list(DATOS_DIR.glob('*.CSV'))
    temas = sorted(list(set([archivo.stem for archivo in archivos])))
    return temas

def load_questions_from_file(archivo):
    preguntas = []
    tema_nombre = archivo.stem
    try:
        if archivo.suffix.lower() == '.xlsx':
            df = pd.read_excel(archivo)
        else:
            for sep in [',', ';', '\t']:
                for encoding in ['utf-8', 'latin-1', 'cp1252']:
                    try:
                        df = pd.read_csv(archivo, sep=sep, encoding=encoding)
                        if len(df.columns) > 1:
                            break
                    except:
                        continue
                else:
                    continue
                break
            else:
                st.warning(f"No se pudo determinar el formato del CSV: {archivo.name}")
                return []
        
        df.columns = [str(col).strip() for col in df.columns]
        column_mapping = {}
        for col in df.columns:
            col_lower = col.lower().replace('ó', 'o').replace('í', 'i').replace('é', 'e')
            if 'pregunta' in col_lower:
                column_mapping['pregunta'] = col
            elif col_lower in ['correcta', 'respuesta', 'correct']:
                column_mapping['correcta'] = col
            elif 'opcion' in col_lower or 'option' in col_lower:
                if col_lower.endswith('a') or 'a)' in col_lower or col_lower == 'opcion a':
                    column_mapping['opcionA'] = col
                elif col_lower.endswith('b') or 'b)' in col_lower or col_lower == 'opcion b':
                    column_mapping['opcionB'] = col
                elif col_lower.endswith('c') or 'c)' in col_lower or col_lower == 'opcion c':
                    column_mapping['opcionC'] = col
            elif 'imagen' in col_lower or 'img' in col_lower:
                column_mapping['imagen'] = col

        opciones_cols = []
        for letra in ['A', 'B', 'C']:
            col_key = f'opcion{letra}'
            if col_key in column_mapping:
                opciones_cols.append(column_mapping[col_key])
            else:
                found = False
                for col in df.columns:
                    col_lower = col.lower().replace('ó', 'o').replace('í', 'i').replace('é', 'e')
                    if letra == 'A' and ('opcion a' in col_lower or col_lower.endswith(' a') or col_lower.endswith('a)') or col_lower == 'a'):
                        opciones_cols.append(col); found = True; break
                    elif letra == 'B' and ('opcion b' in col_lower or col_lower.endswith(' b') or col_lower.endswith('b)') or col_lower == 'b'):
                        opciones_cols.append(col); found = True; break
                    elif letra == 'C' and ('opcion c' in col_lower or col_lower.endswith(' c') or col_lower.endswith('c)') or col_lower == 'c'):
                        opciones_cols.append(col); found = True; break
                if not found:
                    opciones_cols.append(None)

        pregunta_col = column_mapping.get('pregunta')
        correcta_col = column_mapping.get('correcta')
        imagen_col = column_mapping.get('imagen')

        if not pregunta_col:
            st.warning(f"⚠️ El archivo '{archivo.name}' se abrió pero no tiene una columna llamada 'Pregunta'. Columnas encontradas: {list(df.columns)}")
            return []

        for _, fila in df.iterrows():
            if pd.isna(fila[pregunta_col]) or str(fila[pregunta_col]).strip() == '':
                continue
            
            opciones = []
            for i, col in enumerate(opciones_cols):
                if col and col in df.columns and not pd.isna(fila[col]):
                    opciones.append(str(fila[col]).strip())
                else:
                    opciones.append(f'Opción {chr(65+i)} (no definida)')
            
            while len(opciones) < 3:
                opciones.append(f'Opción {chr(65+len(opciones))} (no definida)')
            
            correcta = 'A'
            if correcta_col and correcta_col in df.columns and not pd.isna(fila[correcta_col]):
                correcta_raw = str(fila[correcta_col]).upper().strip()
                if correcta_raw in ['A', 'B', 'C']:
                    correcta = correcta_raw
                elif correcta_raw in ['1', '2', '3']:
                    correcta = chr(64 + int(correcta_raw))
            
            img_path = None
            if imagen_col and imagen_col in df.columns and not pd.isna(fila[imagen_col]):
                img_path = str(fila[imagen_col]).strip()

            preguntas.append({
                'tema': tema_nombre,
                'pregunta': str(fila[pregunta_col]).strip(),
                'opciones': opciones[:3],
                'correcta': correcta,
                'imagen': img_path
            })
    except Exception as e:
        st.error(f"Error cargando el archivo {archivo.name}: {e}")
        
    return preguntas

def load_questions(tema=None):
    if not DATOS_DIR.exists():
        return []
    todas_preguntas = []
    if tema:
        for extension in ['.xlsx', '.csv', '.XLSX', '.CSV']:
            archivo = DATOS_DIR / f"{tema}{extension}"
            if archivo.exists():
                todas_preguntas.extend(load_questions_from_file(archivo))
                break
    else:
        archivos = list(DATOS_DIR.glob('*.xlsx')) + list(DATOS_DIR.glob('*.csv')) + list(DATOS_DIR.glob('*.XLSX')) + list(DATOS_DIR.glob('*.CSV'))
        for archivo in archivos:
            todas_preguntas.extend(load_questions_from_file(archivo))
    return todas_preguntas

# --- ESTADOS DE LA SESIÓN ---
if 'view' not in st.session_state:
    st.session_state.view = 'home'
if 'exam_questions' not in st.session_state:
    st.session_state.exam_questions = []
if 'current_index' not in st.session_state:
    st.session_state.current_index = 0
if 'user_answers' not in st.session_state:
    st.session_state.user_answers = {}
if 'exam_title' not in st.session_state:
    st.session_state.exam_title = ""

# --- NAVEGACIÓN ---
st.title("📚 Sistema de Exámenes Profesional")

with st.sidebar:
    st.header("Menú Navegación")
    if st.button("🏠 Inicio", use_container_width=True):
        st.session_state.view = 'home'
        st.rerun()
    if st.button("📊 Estadísticas", use_container_width=True):
        st.session_state.view = 'stats'
        st.rerun()

# --- VISTA: INICIO Y CONFIGURACIÓN ---
if st.session_state.view == 'home':
    st.subheader("Selecciona el tipo de examen")
    
    opciones_examen = st.radio(
        "Modo de examen:",
        ["📖 Examen por Tema Específico", "📚 Examen Multi-Tema", "🎯 Examen Completo", "❌ Preguntas más Falladas"]
    )
    
    temas = get_available_topics()
    preguntas_seleccionadas = []
    titulo_examen = ""

    if not temas:
        st.error("❌ No se encontraron archivos en la carpeta 'datos/'. Asegúrate de que la carpeta existe en GitHub y contiene archivos Excel (.xlsx) o CSV.")
    else:
        if opciones_examen == "📖 Examen por Tema Específico":
            tema_sel = st.selectbox("Selecciona un tema:", temas)
            if tema_sel:
                preguntas_seleccionadas = load_questions(tema_sel)
                titulo_examen = f"Examen: {tema_sel}"

        elif opciones_examen == "📚 Examen Multi-Tema":
            temas_sel = st.multiselect("Selecciona uno o varios temas:", temas)
            if temas_sel:
                for t in temas_sel:
                    preguntas_seleccionadas.extend(load_questions(t))
                titulo_examen = f"Examen Multi-Tema"

        elif opciones_examen == "🎯 Examen Completo":
            preguntas_seleccionadas = load_questions()
            titulo_examen = "Examen Completo"

        elif opciones_examen == "❌ Preguntas más Falladas":
            stats = load_statistics()
            failed_questions = []
            for key, stat in stats.items():
                if stat['intentos'] >= 2 and stat['fallos'] > 0:
                    failed_questions.append(stat)
            
            failed_questions.sort(key=lambda x: (x['fallos'] / x['intentos']), reverse=True)
            failed_questions = failed_questions[:20]
            
            all_q = load_questions()
            for f in failed_questions:
                for q in all_q:
                    if q['tema'] == f['tema'] and q['pregunta'] == f['pregunta']:
                        preguntas_seleccionadas.append(q)
                        break
            titulo_examen = "Preguntas más Falladas"

        st.divider()
        if preguntas_seleccionadas:
            st.success(f"✅ Preguntas cargadas correctamente: **{len(preguntas_seleccionadas)}**")
            
            col1, col2 = st.columns(2)
            with col1:
                num_q = st.slider("Número de preguntas a realizar:", 1, len(preguntas_seleccionadas), min(10, len(preguntas_seleccionadas)))
            with col2:
                mezclar = st.checkbox("Mezclar preguntas", value=True)

            if st.button("🚀 Iniciar Examen", type="primary", use_container_width=True):
                if mezclar:
                    random.shuffle(preguntas_seleccionadas)
                
                st.session_state.exam_questions = preguntas_seleccionadas[:num_q]
                st.session_state.current_index = 0
                st.session_state.user_answers = {}
                st.session_state.exam_title = titulo_examen
                st.session_state.view = 'exam'
                st.rerun()
        else:
            st.info("💡 Por favor, selecciona las opciones anteriores para cargar las preguntas.")

# --- VISTA: EXAMEN ---
elif st.session_state.view == 'exam':
    idx = st.session_state.current_index
    questions = st.session_state.exam_questions
    q_curr = questions[idx]

    st.subheader(f"{st.session_state.exam_title}")
    st.progress((idx + 1) / len(questions), text=f"Pregunta {idx + 1} de {len(questions)}")

    st.markdown(f"### ❓ {q_curr['pregunta']}")

    if q_curr.get('imagen'):
        raw_img = str(q_curr['imagen']).strip().replace('\\', '/')
        
        # Elimina 'imagenes/' si viene escrito en el Excel
        if raw_img.lower().startswith('imagenes/'):
            nombre_limpio = raw_img[9:]
        else:
            nombre_limpio = raw_img
            
        img_file = IMAGENES_DIR / nombre_limpio
        
        # Búsqueda directa o insensible a mayúsculas
        if img_file.exists():
            st.image(str(img_file), use_container_width=True)
        else:
            encontrada = False
            if IMAGENES_DIR.exists():
                for f in IMAGENES_DIR.iterdir():
                    if f.name.lower() == nombre_limpio.lower():
                        st.image(str(f), use_container_width=True)
                        encontrada = True
                        break
            if not encontrada:
                st.warning(f"⚠️ No se encontró la imagen: '{nombre_limpio}' dentro de la carpeta 'imagenes/'.")

    options_map = {f"{chr(65+i)}) {opt}": chr(65+i) for i, opt in enumerate(q_curr['opciones'])}
    
    saved_ans = st.session_state.user_answers.get(idx, None)
    saved_index = None
    if saved_ans:
        for i, key in enumerate(options_map.keys()):
            if options_map[key] == saved_ans:
                saved_index = i

    selected_option = st.radio(
        "Selecciona tu respuesta:",
        list(options_map.keys()),
        index=saved_index,
        key=f"q_{idx}"
    )

    if selected_option:
        chosen_letter = options_map[selected_option]
        st.session_state.user_answers[idx] = chosen_letter
        
        if chosen_letter == q_curr['correcta']:
            st.success("✅ ¡Correcto!")
        else:
            st.error(f"❌ Incorrecto. La respuesta correcta es la **{q_curr['correcta']}**")

    st.divider()
    col_back, col_next = st.columns(2)
    
    with col_back:
        if idx > 0:
            if st.button("⬅️ Anterior", use_container_width=True):
                st.session_state.current_index -= 1
                st.rerun()

    with col_next:
        if idx < len(questions) - 1:
            if st.button("➡️ Siguiente", use_container_width=True):
                st.session_state.current_index += 1
                st.rerun()
        else:
            if st.button("✅ Finalizar Examen", type="primary", use_container_width=True):
                st.session_state.view = 'results'
                st.rerun()

# --- VISTA: RESULTADOS ---
elif st.session_state.view == 'results':
    st.subheader("📊 Resultados del Examen")
    questions = st.session_state.exam_questions
    answers = st.session_state.user_answers
    
    correct_count = 0
    for i, q in enumerate(questions):
        user_ans = answers.get(i)
        is_corr = (user_ans == q['correcta'])
        if is_corr:
            correct_count += 1
        update_statistics(q, is_corr)

    total = len(questions)
    percentage = (correct_count / total) * 100 if total > 0 else 0

    c1, c2, c3 = st.columns(3)
    c1.metric("Aciertos", f"{correct_count} / {total}")
    c2.metric("Fallos", f"{total - correct_count}")
    c3.metric("Porcentaje", f"{percentage:.1f}%")

    if percentage >= 80:
        st.balloons()
        st.success("🎉 ¡Excelente resultado!")
    elif percentage >= 50:
        st.info("👍 Buen trabajo, pero puedes seguir repasando.")
    else:
        st.error("📖 Necesitas estudiar más este tema.")

    st.divider()
    st.markdown("### Detalle de respuestas")
    
    for i, q in enumerate(questions):
        user_ans = answers.get(i, "Sin responder")
        corr = q['correcta']
        with st.expander(f"Pregunta {i+1}: {'✅ Correcta' if user_ans == corr else '❌ Incorrecta'}"):
            st.write(f"**{q['pregunta']}**")
            st.write(f"- Tu respuesta: {user_ans}")
            st.write(f"- Respuesta correcta: **{corr}**")

    if st.button("🏠 Volver al Inicio", type="primary"):
        st.session_state.view = 'home'
        st.rerun()

# --- VISTA: ESTADÍSTICAS ---
elif st.session_state.view == 'stats':
    st.subheader("📊 Estadísticas acumuladas")
    stats = load_statistics()
    
    if not stats:
        st.info("Aún no hay estadísticas registradas. Completa algunos exámenes para ver los datos aquí.")
    else:
        df = pd.DataFrame(stats.values())
        st.dataframe(df[['tema', 'pregunta', 'intentos', 'aciertos', 'fallos']], use_container_width=True)
