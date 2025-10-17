# App Sicar Productos — Excel ⇄ MySQL (estandarizada)

## Objetivo
- SUBIR usa **2 sentencias SQL** exactas: `SELECT` catálogo `unidad` y `UPDATE` masivo a `articulo`.
- Antes de subir se limpian, validan y convierten datos del Excel.
- `servicio` y `granel` se fuerzan a booleanos `0/1`.
- DESCARGA es un `SELECT` de solo lectura para vista/exportación.

## Requisitos
- Windows 10/11, macOS 12+, o Linux con X11/Wayland.
- Python 3.10–3.12.
- Acceso a MySQL y tabla `articulo`/`unidad` existentes.

## Instalación
```bash
python -m venv .venv
# Windows
.venv\Scripts\pip install -r requirements.txt
# macOS/Linux
source .venv/bin/activate && pip install -r requirements.txt
cp .env.example .env
# Edita .env con tus credenciales
```

## Ejecución
```bash
# Windows
.venv\Scripts\python app_sicar_productos.py
# macOS/Linux
python app_sicar_productos.py
```

## Configuración (.env)
```ini
MYSQL_HOST=127.0.0.1
MYSQL_PORT=3306
MYSQL_USER=root
MYSQL_PASSWORD=changeme
MYSQL_DB=sicar
DATA_FOLDER=/ruta/opcional
APP_LOG_LEVEL=INFO
```

## Flujo
1. **Descargar**: llena el Excel seleccionado con el `SELECT` de productos.
2. **Editar Excel**: respeta encabezados. Debe incluir **id** o **clave**.
3. **Subir**: se limpia y se hace `UPDATE` masivo. Si unidad no existe, se manda `NULL` para no violar FK.

## Columnas soportadas
- `descripcion` `existencia` `servicio` `precio_compra` `precio1` `precio2` `precio3`
- `unidad_compra` `unidad_venta` `factor` `granel` `clave_alterna`
- Identificador obligatorio: `id` o `clave`.

## Reglas de limpieza resumidas
- Numéricos: acepta `12,50` o texto con números. Vacíos → `NULL`.
- Booleanos: `{1,0,true,false,si,no}` → `1/0`. Vacío → `NULL`.
- Unidades: alias comunes (`PZA`,`LT`, etc.) → nombre canónico → `uni_id`. Desconocido → `NULL`.

## Empaquetado opcional (EXE/APP)
```bash
pip install pyinstaller
pyinstaller --noconfirm --onefile --name SicarExcelMySQL app_sicar_productos.py
# Binario en dist/
```

## SQL requeridos (referencia)
- `articulo` con columnas: `art_id`, `clave`, `clavealterna`, `descripcion`, `existencia`,
  `servicio`, `preciocompra`, `precio1..3`, `unidadCompra`, `unidadVenta`, `factor`, `granel`, `status`.
- `unidad` con `uni_id`, `nombre`.

## Seguridad
- No hay passwords en código. Usa `.env` o variables de entorno del sistema.
- Recomendado crear un usuario MySQL con permisos mínimos para `SELECT` y `UPDATE` en estas tablas.
