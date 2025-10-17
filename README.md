# Manual Técnico Avanzado — Aplicación **Sicar Productos 2SQL** (Interfaz Excel ⇄ MySQL)

## 1. Introducción y contexto conceptual
La aplicación **Sicar Productos 2SQL** constituye un sistema de intermediación de datos diseñado para sincronizar información entre hojas de cálculo **Microsoft Excel** y bases de datos **MySQL** en entornos empresariales que operan bajo la arquitectura del sistema **SICAR**. Concebida como una herramienta de integración heterogénea, la aplicación fue desarrollada en **Python** empleando **PyQt5** como capa de interfaz gráfica y un conjunto modular de bibliotecas de manipulación y persistencia de datos.

Desde una perspectiva arquitectónica, el sistema persigue la estandarización de flujos bidireccionales de información, garantizando consistencia semántica, tipificación estricta y trazabilidad total del proceso de actualización. La aplicación abstrae las interacciones de bajo nivel con la base de datos y ofrece una interfaz visual que facilita operaciones complejas a usuarios técnicos y no técnicos.

## 2. Requerimientos técnicos y entorno operativo
La infraestructura tecnológica de **Sicar Productos 2SQL** se fundamenta en los siguientes componentes:

- **Lenguaje base:** Python ≥ 3.9, seleccionado por su ecosistema científico y su compatibilidad multiplataforma.
- **Motor de base de datos:** MySQL ≥ 5.7, con soporte de codificación UTF-8MB4 y estructura conforme al esquema `sicar`.
- **Entorno gráfico:** PyQt5, optimizado para renderizado de alta densidad (High DPI) y uso extensivo de widgets nativos.
- **Dependencias complementarias:**
  - `numpy` y `pandas` para manipulación vectorial y tabular.
  - `python-dotenv` para la inyección controlada de variables de entorno.
  - `openpyxl` para la serialización de datos estructurados en formato `.xlsx`.
  - `pymysql` y opcionalmente `SQLAlchemy` para conectividad transaccional.
- **Sistemas soportados:** Microsoft Windows, macOS y distribuciones Linux con acceso local o remoto a MySQL.

## 3. Instalación y validación del entorno
### 3.1 Procedimiento de despliegue
1. Crear entorno virtual:
   ```bash
   python -m venv .venv
   ```
2. Activar entorno según sistema operativo:
   ```bash
   source .venv/bin/activate   # Linux/macOS
   .venv\Scripts\activate    # Windows
   ```
3. Instalar dependencias:
   ```bash
   pip install -r requirements.txt
   ```
4. Verificar instalación de PyQt5 y conectividad MySQL:
   ```bash
   python -m PyQt5.QtCore
   ```

### 3.2 Configuración de entorno
El archivo `.env` define parámetros de conexión y de comportamiento de la aplicación:
```
MYSQL_HOST=127.0.0.1
MYSQL_PORT=3306
MYSQL_USER=root
MYSQL_PASSWORD=clave_segura
MYSQL_DB=sicar
APP_LOG_LEVEL=INFO
DATA_FOLDER=~/Documents/SicarData
MAX_PREVIEW_ROWS=50000
```
Cada variable puede redefinirse mediante entorno del sistema, permitiendo integraciones CI/CD o automatización de despliegues.

## 4. Modelo arquitectónico
La aplicación se organiza bajo un modelo **multicapa** con separación estricta de responsabilidades:
1. **Capa de presentación:** Implementada en **PyQt5**, gestiona interacción visual, renderizado de tablas, filtrado dinámico y ejecución de comandos de usuario.
2. **Capa lógica de negocio:** Traduce las acciones de la interfaz en operaciones SQL y rutinas de validación, incluyendo control transaccional y tipificación coercitiva.
3. **Capa de persistencia:** Establece conexiones seguras con MySQL mediante `pymysql` o `SQLAlchemy`, administrando commits y rollbacks según el resultado operativo.

### 4.1 Principales componentes
- **Clase `App(QWidget)`:** núcleo de control de interfaz y orquestación de eventos.
- **Módulo de E/S tabular:** lectura y escritura robusta de archivos Excel vía `pandas` y `openpyxl`.
- **Subsistema SQL:** definición de consultas `SQL_PRODUCTOS` y `SQL_LOOKUPS` para sincronización estructurada.
- **Validador de datos:** mecanismos de coerción y normalización que garantizan integridad de tipos.
- **Sistema de logging:** seguimiento de eventos a nivel INFO, WARNING y ERROR para auditoría.

### 4.2 Diagrama de flujo lógico
```
Excel (.xlsx) ↔ Interfaz PyQt5 ↔ Validación estructural ↔ Generación de SQL ↔ MySQL
                                          ↑                                        ↓
                               Búsqueda, filtrado y vista previa      Actualización y retroalimentación
```

## 5. Ciclo funcional
### 5.1 Descarga de base de datos (DB → Excel)
- El usuario inicia el proceso desde la interfaz principal.
- El sistema ejecuta la consulta `SQL_PRODUCTOS` para extraer el inventario activo.
- Los resultados se exportan a Excel aplicando formatos de encabezado, filtros y ancho de columnas automático.
- La vista previa de datos se sincroniza con el archivo resultante.

### 5.2 Subida de datos (Excel → DB)
- Se verifica la existencia de una columna clave (`id` o `clave`).
- Se ejecuta una limpieza semántica del DataFrame, eliminando registros inconsistentes.
- Se consulta `SQL_LOOKUPS` para obtener identificadores normalizados de unidades.
- Se aplica conversión de tipos numéricos, booleanos y textuales.
- Se construye un comando **UPDATE** parametrizado que actualiza múltiples registros mediante `executemany()`.
- El proceso culmina con una transacción confirmada y un reporte detallado de filas afectadas y descartadas.

## 6. Tipificación, validación y normalización
Los procedimientos de validación implementan una taxonomía robusta:
- **Campos numéricos:** transformados mediante expresiones regulares que detectan patrones flotantes y enteros.
- **Campos booleanos:** se aceptan valores en diversas representaciones culturales (`sí`, `true`, `1`, `no`, `false`, `0`).
- **Campos de unidad:** se someten a normalización morfológica conforme al diccionario `UNIT_NORMALIZATION`.
- **Cadenas textuales:** limpiadas de espacios, puntuaciones y vacíos lógicos.
- **Valores nulos:** representados como `NULL` para preservar integridad referencial.

## 7. Manejo de errores y trazabilidad
La aplicación incorpora una estrategia de gestión de excepciones basada en rollback automático y logging persistente. Cada error genera una entrada con contexto temporal, tipo de excepción y trazado del módulo origen.

Errores críticos controlados:
- Fallas de conectividad con MySQL.
- Archivos Excel corruptos o con encabezados no conformes.
- Violaciones de tipo o de integridad de datos.
- Ausencia de dependencias esenciales.

Los errores de usuario se comunican mediante cuadros de diálogo (`QMessageBox`) y notificaciones en la barra de estado.

## 8. Rendimiento y optimización
El rendimiento del sistema se sustenta en:
- Previsualización limitada por `MAX_PREVIEW_ROWS` para optimizar carga en memoria.
- Uso de operaciones vectorizadas de `pandas` y `numpy`.
- Ejecución transaccional masiva mediante `executemany()`.
- Reutilización de conexiones persistentes con `pool_pre_ping` y `pool_recycle` en `SQLAlchemy`.

Estas medidas permiten procesar decenas de miles de registros en operaciones de actualización sin degradación significativa del rendimiento.

## 9. Seguridad y cumplimiento
El diseño evita la inclusión de credenciales embebidas y privilegios excesivos. Las credenciales se leen desde variables de entorno y las conexiones se limitan a usuarios con permisos de actualización, no de modificación estructural. El sistema no expone servicios de red ni interfaces HTTP, garantizando aislamiento operativo.

## 10. Portabilidad y mantenimiento
El código evita rutas absolutas, utilizando `os.path.expanduser('~')` y mecanismos de autodetección de directorio de datos. La interfaz se adapta a entornos oscuros y pantallas HiDPI. Para mantenimiento:
- Ejecutar `pip install -U -r requirements.txt` periódicamente.
- Verificar coherencia entre columnas SQL y mapeo `ALLOWED_MAP`.
- Registrar versiones del esquema `sicar` para trazabilidad histórica.

## 11. Despliegue y distribución
### 11.1 Distribución en código fuente
El proyecto puede distribuirse junto con `requirements.txt` y un `.env` de referencia.

### 11.2 Generación de binario
```bash
pyinstaller -y --name Sicar2SQL --windowed --icon icon.ico main.py
```
El ejecutable empaquetado puede incluirse con la carpeta `SicarData` preconfigurada para entornos de producción.

## 12. Evaluación, pruebas y aseguramiento de calidad
Se recomienda un protocolo de validación basado en cinco etapas:
1. **Inicio:** verificación de carga de pestañas y recursos.
2. **Descarga:** comparación de resultados entre MySQL y Excel.
3. **Actualización:** validación cruzada de registros modificados.
4. **Filtro:** análisis de rendimiento en búsquedas textuales.
5. **Errores:** simulación de fallos controlados para testear robustez.

## 13. Conclusión
**Sicar Productos 2SQL** representa una implementación madura y extensible de interoperabilidad entre Excel y MySQL. Su enfoque modular, su control de calidad de datos y su neutralidad multiplataforma la convierten en una solución idónea para entornos corporativos que requieren consistencia transaccional, reproducibilidad y trazabilidad en la administración de inventarios y catálogos técnicos.
