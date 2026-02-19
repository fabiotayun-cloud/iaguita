# Reconocimiento Facial - Asistencias

Aplicacion web para registrar asistencias a entrenamientos de futbol mediante reconocimiento facial, y tambien comparar rostros individualmente.

## Instalacion

1. Crear entorno virtual e instalar dependencias:
```bash
python -m venv venv
venv\Scripts\activate
pip install -r requirements.txt
```

> En Windows con Python 3.12+, `dlib-bin` se instala pre-compilado. Si `face_recognition` intenta compilar `dlib`, instalar con: `pip install dlib-bin` y luego `pip install face_recognition --no-deps`.

## Configurar miembros

Colocar una foto de cada miembro del curso en la carpeta `miembros/`. El nombre del archivo (sin extension) sera el nombre del miembro:

```
miembros/
  Juan Perez.jpg
  Maria Lopez.png
  Carlos Ruiz.jpeg
```

Requisitos de las fotos:
- Una sola persona por foto (selfie o foto clara del rostro)
- Formatos: jpg, jpeg, png, webp, bmp

## Uso

1. Ejecutar la aplicacion:
```bash
python app.py
```

2. Abrir `http://localhost:5000`

### Pestaña: Asistencias

1. Verificar que los miembros estan cargados (indicador en la barra superior)
2. Subir la foto grupal del entrenamiento
3. La app extrae la fecha automaticamente desde los metadatos EXIF de la foto
4. Click en "Registrar Asistencia"
5. Se muestra quien asistio y quien no, y se actualiza el Excel `asistencias.xlsx`
6. Descargar el Excel con el boton "Descargar Excel"

Si se agregan nuevos miembros a la carpeta, usar el boton "Recargar Miembros".

### Pestaña: Comparar Rostros

Funcionalidad original: subir una foto de referencia y una foto objetivo para verificar si apareces en ella.

## Estructura del proyecto

```
Reconocimiento facial/
  app.py              # Backend Flask
  requirements.txt    # Dependencias
  templates/
    index.html        # Frontend
  miembros/           # Fotos de referencia de cada miembro
  asistencias.xlsx    # Excel de asistencias (generado automaticamente)
```

## Excel de asistencias

El archivo `asistencias.xlsx` se genera automaticamente con el formato:

| Miembro      | 2026-02-10 | 2026-02-12 | ... |
|--------------|------------|------------|-----|
| Juan Perez   | ✓          |            |     |
| Maria Lopez  | ✓          | ✓          |     |
