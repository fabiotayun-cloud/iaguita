import os
import re
import traceback
from datetime import datetime
from pathlib import Path
from flask import Flask, render_template, request, jsonify, send_file
import face_recognition
from PIL import Image
from PIL.ExifTags import TAGS
import numpy as np
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024  # 32MB max

BASE_DIR = Path(__file__).parent
MEMBERS_DIR = BASE_DIR / 'miembros'
EXCEL_PATH = BASE_DIR / 'asistencias.xlsx'
ALLOWED_EXTENSIONS = {'png', 'jpg', 'jpeg', 'webp', 'bmp'}

# Cache global de miembros: {nombre: encoding}
members_cache = {}


# ─────────────────────────────────────────────
# Utilidades generales
# ─────────────────────────────────────────────

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def load_image_from_upload(file_storage):
    """Carga una imagen desde un FileStorage de Flask y la convierte a formato RGB numpy array."""
    img = Image.open(file_storage)
    img = img.convert('RGB')
    return np.array(img)


def detect_faces_multiscale(image, upsample=2):
    """
    Detecta rostros usando HOG con múltiples escalas para encontrar caras pequeñas.
    Escala la imagen hacia arriba en varias pasadas y combina resultados,
    eliminando detecciones duplicadas.
    """
    all_locations = set()

    # Pasada 1: imagen original con upsample configurado
    locs = face_recognition.face_locations(image, number_of_times_to_upsample=upsample, model='hog')
    for loc in locs:
        all_locations.add(loc)

    # Pasada 2: imagen escalada x1.5 (para caras pequeñas)
    h, w = image.shape[:2]
    scale = 1.5
    scaled = np.array(Image.fromarray(image).resize((int(w * scale), int(h * scale)), Image.LANCZOS))
    locs_scaled = face_recognition.face_locations(scaled, number_of_times_to_upsample=1, model='hog')
    for (top, right, bottom, left) in locs_scaled:
        orig_loc = (
            int(top / scale),
            int(right / scale),
            int(bottom / scale),
            int(left / scale)
        )
        if not _is_duplicate(orig_loc, all_locations):
            all_locations.add(orig_loc)

    # Pasada 3: imagen escalada x2.0 (para caras muy pequeñas)
    if upsample >= 2:
        scale2 = 2.0
        scaled2 = np.array(Image.fromarray(image).resize((int(w * scale2), int(h * scale2)), Image.LANCZOS))
        locs_scaled2 = face_recognition.face_locations(scaled2, number_of_times_to_upsample=1, model='hog')
        for (top, right, bottom, left) in locs_scaled2:
            orig_loc = (
                int(top / scale2),
                int(right / scale2),
                int(bottom / scale2),
                int(left / scale2)
            )
            if not _is_duplicate(orig_loc, all_locations):
                all_locations.add(orig_loc)

    return list(all_locations)


def _is_duplicate(new_loc, existing_locations, iou_threshold=0.4):
    """Verifica si una detección ya existe (basado en IoU - Intersection over Union)."""
    new_top, new_right, new_bottom, new_left = new_loc
    for (top, right, bottom, left) in existing_locations:
        inter_top = max(new_top, top)
        inter_left = max(new_left, left)
        inter_bottom = min(new_bottom, bottom)
        inter_right = min(new_right, right)

        if inter_bottom <= inter_top or inter_right <= inter_left:
            continue

        inter_area = (inter_bottom - inter_top) * (inter_right - inter_left)
        area1 = (new_bottom - new_top) * (new_right - new_left)
        area2 = (bottom - top) * (right - left)
        union_area = area1 + area2 - inter_area

        if union_area > 0 and inter_area / union_area > iou_threshold:
            return True
    return False


# ─────────────────────────────────────────────
# Gestión de miembros
# ─────────────────────────────────────────────

def load_members():
    """Escanea la carpeta miembros/ y calcula encodings faciales para cada uno."""
    global members_cache
    members_cache = {}

    if not MEMBERS_DIR.exists():
        MEMBERS_DIR.mkdir(parents=True, exist_ok=True)
        return

    loaded = 0
    errors = []
    for file_path in MEMBERS_DIR.iterdir():
        if not file_path.is_file():
            continue
        if file_path.suffix.lower().lstrip('.') not in ALLOWED_EXTENSIONS:
            continue

        name = file_path.stem  # nombre sin extensión
        try:
            img = Image.open(file_path).convert('RGB')
            img_array = np.array(img)
            encodings = face_recognition.face_encodings(img_array)
            if len(encodings) > 0:
                members_cache[name] = encodings[0]
                loaded += 1
            else:
                errors.append(f'{name}: no se detectó rostro')
        except Exception as e:
            errors.append(f'{name}: {str(e)}')

    print(f'[Miembros] Cargados: {loaded} | Errores: {len(errors)}')
    for err in errors:
        print(f'  - {err}')


# ─────────────────────────────────────────────
# Extracción de metadatos EXIF
# ─────────────────────────────────────────────

def extract_photo_date(file_storage):
    """Extrae la fecha de captura: primero EXIF, luego nombre del archivo.
    Retorna (fecha_str, fuente) donde fuente es 'exif', 'nombre' o None."""
    # 1. Intentar EXIF
    try:
        file_storage.seek(0)
        img = Image.open(file_storage)

        exif_data = None
        if hasattr(img, '_getexif') and img._getexif():
            exif_data = img._getexif()

        if not exif_data:
            exif_obj = img.getexif()
            if exif_obj:
                exif_data = dict(exif_obj)

        if exif_data:
            date_value = None
            for tag_id in [36867, 36868, 306]:
                if tag_id in exif_data and exif_data[tag_id]:
                    date_value = exif_data[tag_id]
                    break

            if date_value and isinstance(date_value, str):
                for fmt in ['%Y:%m:%d %H:%M:%S', '%Y-%m-%d %H:%M:%S', '%Y:%m:%d']:
                    try:
                        dt = datetime.strptime(date_value.strip(), fmt)
                        file_storage.seek(0)
                        print(f'[Fecha] Encontrada en EXIF: {dt.strftime("%Y-%m-%d")}')
                        return dt.strftime('%Y-%m-%d'), 'exif'
                    except ValueError:
                        continue

    except Exception as e:
        print(f'[Fecha] Error extrayendo EXIF: {e}')

    # 2. Intentar extraer fecha del nombre del archivo
    file_storage.seek(0)
    filename = file_storage.filename or ''
    date_from_name = extract_date_from_filename(filename)
    if date_from_name:
        print(f'[Fecha] Encontrada en nombre de archivo "{filename}": {date_from_name}')
        return date_from_name, 'nombre'

    file_storage.seek(0)
    return None, None


def extract_date_from_filename(filename):
    """Extrae una fecha del nombre de archivo. Soporta formatos comunes como
    'WhatsApp Image 2026-02-05 at 8.43.17 PM', 'IMG_20260205_...', '2026-02-05 foto', etc."""
    if not filename:
        return None

    # Patrón: YYYY-MM-DD
    match = re.search(r'(\d{4})-(\d{2})-(\d{2})', filename)
    if match:
        try:
            dt = datetime(int(match.group(1)), int(match.group(2)), int(match.group(3)))
            return dt.strftime('%Y-%m-%d')
        except ValueError:
            pass

    # Patrón: YYYYMMDD (ej: IMG_20260205_)
    match = re.search(r'(\d{4})(\d{2})(\d{2})', filename)
    if match:
        try:
            dt = datetime(int(match.group(1)), int(match.group(2)), int(match.group(3)))
            if 2000 <= dt.year <= 2100:
                return dt.strftime('%Y-%m-%d')
        except ValueError:
            pass

    # Patrón: DD-MM-YYYY o DD/MM/YYYY
    match = re.search(r'(\d{2})[-/](\d{2})[-/](\d{4})', filename)
    if match:
        try:
            dt = datetime(int(match.group(3)), int(match.group(2)), int(match.group(1)))
            return dt.strftime('%Y-%m-%d')
        except ValueError:
            pass

    return None


# ─────────────────────────────────────────────
# Gestión del Excel de asistencias
# ─────────────────────────────────────────────

def load_or_create_excel():
    """Carga el Excel existente o crea uno nuevo. Retorna (workbook, worksheet)."""
    if EXCEL_PATH.exists():
        wb = load_workbook(str(EXCEL_PATH))
        ws = wb.active
    else:
        wb = Workbook()
        ws = wb.active
        ws.title = 'Asistencias'
        ws['A1'] = 'Miembro'
        _style_header(ws, 1, 1)
    return wb, ws


def _style_header(ws, row, col):
    """Aplica estilo a una celda de encabezado."""
    cell = ws.cell(row=row, column=col)
    cell.font = Font(bold=True, color='FFFFFF')
    cell.fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
    cell.alignment = Alignment(horizontal='center', vertical='center')


def update_attendance_excel(date_str, present_members):
    """
    Actualiza el Excel de asistencias.
    - date_str: fecha del entrenamiento (ej: '2026-02-18')
    - present_members: lista de nombres de miembros presentes
    """
    wb, ws = load_or_create_excel()

    # Obtener todas las fechas existentes (fila 1, desde columna 2)
    date_columns = {}
    for col in range(2, ws.max_column + 1):
        val = ws.cell(row=1, column=col).value
        if val:
            date_columns[str(val)] = col

    # Si la fecha no existe, agregar nueva columna
    if date_str not in date_columns:
        new_col = ws.max_column + 1 if ws.max_column >= 2 else 2
        ws.cell(row=1, column=new_col, value=date_str)
        _style_header(ws, 1, new_col)
        ws.column_dimensions[ws.cell(row=1, column=new_col).column_letter].width = 14
        date_columns[date_str] = new_col

    date_col = date_columns[date_str]

    # Obtener miembros existentes en el Excel (columna A, desde fila 2)
    member_rows = {}
    for row in range(2, ws.max_row + 1):
        val = ws.cell(row=row, column=1).value
        if val:
            member_rows[val] = row

    # Agregar miembros nuevos que no estén en el Excel
    all_members = sorted(set(list(members_cache.keys()) + list(member_rows.keys())))
    for member_name in all_members:
        if member_name not in member_rows:
            new_row = ws.max_row + 1 if ws.max_row >= 2 else 2
            ws.cell(row=new_row, column=1, value=member_name)
            ws.cell(row=new_row, column=1).font = Font(bold=True)
            member_rows[member_name] = new_row

    # Marcar asistencia
    for member_name in all_members:
        row = member_rows[member_name]
        if member_name in present_members:
            cell = ws.cell(row=row, column=date_col, value='✓')
            cell.font = Font(color='008000', bold=True, size=14)
            cell.alignment = Alignment(horizontal='center')
        else:
            existing = ws.cell(row=row, column=date_col).value
            if not existing:
                ws.cell(row=row, column=date_col, value='')

    # Ajustar ancho de columna A
    ws.column_dimensions['A'].width = max(20, max((len(n) + 2 for n in all_members), default=20))

    # Guardar
    wb.save(str(EXCEL_PATH))
    return True


def get_attendance_summary():
    """Lee el Excel y retorna un resumen de asistencias."""
    if not EXCEL_PATH.exists():
        return {'dates': [], 'members': []}

    wb = load_workbook(str(EXCEL_PATH))
    ws = wb.active

    dates = []
    for col in range(2, ws.max_column + 1):
        val = ws.cell(row=1, column=col).value
        if val:
            dates.append(str(val))

    members = []
    for row in range(2, ws.max_row + 1):
        name = ws.cell(row=row, column=1).value
        if not name:
            continue
        attendance = {}
        total = 0
        for col_idx, date in enumerate(dates, start=2):
            val = ws.cell(row=row, column=col_idx).value
            present = val == '✓'
            attendance[date] = present
            if present:
                total += 1
        members.append({
            'name': name,
            'attendance': attendance,
            'total': total
        })

    return {'dates': dates, 'members': members}


# ─────────────────────────────────────────────
# Rutas
# ─────────────────────────────────────────────

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/members', methods=['GET'])
def get_members():
    """Retorna la lista de miembros cargados."""
    return jsonify({
        'count': len(members_cache),
        'members': sorted(members_cache.keys())
    })


@app.route('/reload-members', methods=['POST'])
def reload_members():
    """Recarga la base de miembros desde la carpeta."""
    load_members()
    return jsonify({
        'count': len(members_cache),
        'members': sorted(members_cache.keys()),
        'message': f'Se cargaron {len(members_cache)} miembros.'
    })


@app.route('/attendance', methods=['POST'])
def register_attendance():
    """Procesa una o varias fotos grupales y registra asistencia."""
    photo_files = request.files.getlist('photos')
    if not photo_files or all(not f.filename for f in photo_files):
        return jsonify({'error': 'Se requiere al menos una foto grupal.'}), 400

    # Filtrar archivos válidos
    valid_files = [f for f in photo_files if f.filename and allowed_file(f.filename)]
    if not valid_files:
        return jsonify({'error': 'Formato de archivo no permitido.'}), 400

    if len(members_cache) == 0:
        return jsonify({'error': 'No hay miembros cargados. Agrega fotos a la carpeta miembros/ y recarga.'}), 400

    # Extraer fecha de la primera foto (EXIF o nombre de archivo)
    auto_date, auto_source = extract_photo_date(valid_files[0])
    manual_date = request.form.get('date', '').strip()
    if manual_date:
        date_str = manual_date
        date_source = 'manual'
    elif auto_date:
        date_str = auto_date
        date_source = auto_source  # 'exif' o 'nombre'
    else:
        date_str = datetime.now().strftime('%Y-%m-%d')
        date_source = 'hoy'

    # Parámetros
    upsample = int(request.form.get('upsample', 2))
    tolerance = float(request.form.get('tolerance', 0.75))
    use_multiscale = request.form.get('model', 'hog') == 'multiscale'

    try:
        # Recopilar todos los encodings de todas las fotos
        all_face_encodings = []
        total_faces = 0

        for photo_file in valid_files:
            photo_file.seek(0)
            photo_image = load_image_from_upload(photo_file)

            if use_multiscale:
                face_locations = detect_faces_multiscale(photo_image, upsample=upsample)
            else:
                face_locations = face_recognition.face_locations(photo_image, number_of_times_to_upsample=upsample, model='hog')

            encodings = face_recognition.face_encodings(photo_image, face_locations)
            all_face_encodings.extend(encodings)
            total_faces += len(encodings)

        if total_faces == 0:
            return jsonify({
                'error': f'No se detectaron rostros en {len(valid_files)} foto(s).',
                'date': date_str,
                'date_source': date_source
            }), 400

        # Comparar cada miembro contra todos los rostros de todas las fotos
        member_names = sorted(members_cache.keys())
        member_encodings = [members_cache[name] for name in member_names]

        present_members = []
        member_results = []

        for name, member_enc in zip(member_names, member_encodings):
            distances = face_recognition.face_distance(all_face_encodings, member_enc)
            best_idx = int(np.argmin(distances))
            best_distance = float(distances[best_idx])
            confidence = round((1 - best_distance) * 100, 1)
            is_present = bool(best_distance <= tolerance)

            if is_present:
                present_members.append(name)

            member_results.append({
                'name': name,
                'present': is_present,
                'confidence': confidence
            })

        # Actualizar Excel
        update_attendance_excel(date_str, present_members)

        return jsonify({
            'date': date_str,
            'date_source': date_source,
            'photos_processed': len(valid_files),
            'total_faces_detected': total_faces,
            'total_members': len(member_names),
            'present_count': len(present_members),
            'absent_count': len(member_names) - len(present_members),
            'results': member_results
        })

    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': f'Error durante el análisis: {str(e)}'}), 500


@app.route('/attendance-summary', methods=['GET'])
def attendance_summary():
    """Retorna el resumen completo de asistencias."""
    try:
        summary = get_attendance_summary()
        return jsonify(summary)
    except Exception as e:
        return jsonify({'error': str(e)}), 500


@app.route('/delete-attendance', methods=['POST'])
def delete_attendance():
    """Elimina una columna de fecha del Excel de asistencias."""
    date_str = request.json.get('date', '').strip() if request.is_json else ''
    if not date_str:
        return jsonify({'error': 'Se requiere una fecha para eliminar.'}), 400

    if not EXCEL_PATH.exists():
        return jsonify({'error': 'No hay archivo de asistencias.'}), 404

    try:
        wb = load_workbook(str(EXCEL_PATH))
        ws = wb.active

        # Buscar la columna con esa fecha
        target_col = None
        for col in range(2, ws.max_column + 1):
            val = ws.cell(row=1, column=col).value
            if str(val) == date_str:
                target_col = col
                break

        if not target_col:
            return jsonify({'error': f'No se encontró la fecha {date_str} en el registro.'}), 404

        ws.delete_cols(target_col)
        wb.save(str(EXCEL_PATH))

        return jsonify({'message': f'Asistencia del {date_str} eliminada correctamente.'})
    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': f'Error al eliminar: {str(e)}'}), 500


@app.route('/download-excel', methods=['GET'])
def download_excel():
    """Descarga el archivo Excel de asistencias."""
    if not EXCEL_PATH.exists():
        return jsonify({'error': 'No hay archivo de asistencias aún.'}), 404
    return send_file(str(EXCEL_PATH), as_attachment=True, download_name='asistencias.xlsx')


@app.route('/compare', methods=['POST'])
def compare_faces():
    if 'reference' not in request.files or 'target' not in request.files:
        return jsonify({'error': 'Se requieren ambas imágenes: referencia y objetivo.'}), 400

    ref_file = request.files['reference']
    target_file = request.files['target']

    if not ref_file.filename or not target_file.filename:
        return jsonify({'error': 'No se seleccionaron archivos.'}), 400

    if not allowed_file(ref_file.filename) or not allowed_file(target_file.filename):
        return jsonify({'error': 'Formato de archivo no permitido. Use: png, jpg, jpeg, webp, bmp.'}), 400

    try:
        ref_image = load_image_from_upload(ref_file)
        target_image = load_image_from_upload(target_file)
    except Exception as e:
        return jsonify({'error': f'Error al cargar las imágenes: {str(e)}'}), 400

    # Parámetros del frontend
    upsample = int(request.form.get('upsample', 2))
    tolerance = float(request.form.get('tolerance', 0.6))
    use_multiscale = request.form.get('model', 'hog') == 'multiscale'

    try:
        # Obtener encoding de la foto de referencia
        ref_locations = face_recognition.face_locations(ref_image, number_of_times_to_upsample=upsample, model='hog')
        ref_encodings = face_recognition.face_encodings(ref_image, ref_locations)
        if len(ref_encodings) == 0:
            return jsonify({
                'error': 'No se detectó ningún rostro en la foto de referencia. Sube una foto donde se vea claramente tu cara.'
            }), 400

        ref_encoding = ref_encodings[0]

        # Detectar caras en la foto objetivo
        if use_multiscale:
            target_locations = detect_faces_multiscale(target_image, upsample=upsample)
        else:
            target_locations = face_recognition.face_locations(target_image, number_of_times_to_upsample=upsample, model='hog')

        target_encodings = face_recognition.face_encodings(target_image, target_locations)

        if len(target_encodings) == 0:
            return jsonify({
                'found': False,
                'message': 'No se detectaron rostros en la foto objetivo.',
                'total_faces': 0,
                'matches': []
            })

        distances = face_recognition.face_distance(target_encodings, ref_encoding)

        # Solo marcar como match el rostro con menor distancia (mayor confianza)
        best_idx = int(np.argmin(distances))
        best_distance = distances[best_idx]
        best_is_match = bool(best_distance <= tolerance)

        matches = []
        for i, (distance, location) in enumerate(zip(distances, target_locations)):
            top, right, bottom, left = location
            confidence = round((1 - distance) * 100, 1)
            matches.append({
                'face_index': i + 1,
                'is_match': bool(i == best_idx and best_is_match),
                'confidence': confidence,
                'location': {
                    'top': int(top),
                    'right': int(right),
                    'bottom': int(bottom),
                    'left': int(left)
                }
            })

        found = best_is_match
        best_match = matches[best_idx] if found else None

        if found:
            message = f'¡Te encontré! Se detectó tu rostro con {best_match["confidence"]}% de confianza.'
        else:
            message = 'No se encontró tu rostro en la foto objetivo.'

        return jsonify({
            'found': found,
            'message': message,
            'total_faces': len(target_encodings),
            'matches': matches
        })

    except Exception as e:
        traceback.print_exc()
        return jsonify({'error': f'Error durante el análisis: {str(e)}'}), 500


# ─────────────────────────────────────────────
# Inicialización
# ─────────────────────────────────────────────

# Cargar miembros al iniciar
load_members()

if __name__ == '__main__':
    app.run(debug=True, port=5000)
