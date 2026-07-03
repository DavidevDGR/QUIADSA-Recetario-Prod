# Control de Fabricación en Cuba (pantalla táctil)

Aplicación en Python (Tkinter) para gestionar la ejecución de Órdenes de
Fabricación (OF) en una cuba, con 3 fases: identificación/OF, ejecución de
secciones con temporizador, y cierre con peso final + exportación a Excel.

## Instalación

```bash
pip install openpyxl
python main.py
```

Solo depende de `openpyxl` (Tkinter y sqlite3 vienen incluidos en Python).
Pulsa **Esc** para salir del modo pantalla completa mientras desarrollas.

## Estructura

- `config.py` — parámetros ajustables (rutas, márgenes de tiempo, conexión Solmicro).
- `db.py` — acceso a datos:
  - `SolmicroRepository`: lectura de la OF y de los "Tipos de Ruta" en Solmicro.
    **Actualmente en modo MOCK** (datos de ejemplo) para poder probar la app
    sin la conexión real. Para conectar a Solmicro:
    1. Cambia `MOCK_MODE = False` en `db.py`.
    2. Instala el driver correspondiente (p. ej. `pip install pyodbc`).
    3. Completa `SOLMICRO_CONN_STRING` en `config.py`.
    4. Descomenta y adapta las consultas SQL de ejemplo en
       `obtener_orden_fabricacion()` y `_obtener_secciones()` a los nombres
       reales de tablas/columnas de tu Solmicro.
  - `LocalRepository`: base de datos SQLite local donde se registra toda la
    ejecución (operarios, secciones, tiempos, incidencias). Es la fuente de
    datos para la exportación a Excel.
- `export.py` — genera/actualiza `exports/registro_fabricacion.xlsx` con una
  fila por sección, incluyendo: fecha de exportación, estado (CANCELADA/
  FINALIZADA), OF, máquina, operarios, fechas de inicio/fin, tiempo total,
  y tiempo previsto/real e incidencia de cada sección.
  - En producción, `exports/` debería ser una carpeta sincronizada con la
    nube (OneDrive/SharePoint/Google Drive) asociada a
    `fabricacion@quiadsa.com`, o bien puedes activar la subida directa vía
    API (hay un ejemplo con Microsoft Graph comentado al final del fichero).
- `main.py` — interfaz gráfica con las 3 fases.
- `onscreen_keyboard.py` — teclado virtual QWERTY en pantalla, reutilizable en
  cualquier campo de texto (`attach_keyboard(entry)`) o numérico
  (`attach_keyboard(entry, numeric_only=True)`). Se llama así (y no
  `keyboard.py`) para no chocar con el paquete de PyPI `keyboard`.

## Probar la aplicación sin conexión a Solmicro (modo MOCK)

Mientras `MOCK_MODE = True` en `db.py`, puedes escribir **cualquier texto no
vacío** como código de OF (p. ej. `OF-2026-001`, `PRUEBA`, `12345`) y la
aplicación devolverá siempre la misma orden de ejemplo: artículo
`FORMULA-DEMO-A`, 500 kg, centro `CUBA-01`, con 5 secciones de prueba
(distintos tiempos previstos para poder probar el margen del 25%). No hay
que "acertar" ningún código concreto: el mock no valida nada, solo simula
que Solmicro ha respondido.

## Comportamiento clave implementado

- **Fase 1**: alta de uno o varios operarios (Enter o botón Añadir). Los
  operarios añadidos aparecen en una lista donde se pueden seleccionar y
  **eliminar** con el botón correspondiente por si se han equivocado.
  Lectura de OF por código de barras (el lector actúa como teclado + Enter),
  cuba y peso de tara. Todos los campos de texto muestran el **teclado
  virtual QWERTY** en pantalla al pulsarlos.
- **Fase 2**: instrucción de la sección, temporizador ascendente, botones
  Cancelar (arriba-izquierda), Anterior y Siguiente/Finalizar. Si se pulsa
  Siguiente con un desvío de más del 25% (antes o después) respecto al
  tiempo previsto de la sección, se exige un motivo antes de continuar (el
  cuadro de motivo también muestra el teclado virtual).
- **Fase 3**: tiempo total, peso tara (recordado de la Fase 1), campo de
  peso total (con teclado numérico virtual) con cálculo en vivo del
  **peso neto** (peso total − peso tara). Botón Fin que registra, exporta a
  Excel (incluyendo peso tara, peso total y peso neto) y vuelve a la Fase 1.

## Notas de adaptación pendientes

- El lector de código de barras se ha modelado como entrada de teclado
  (comportamiento estándar de la mayoría de lectores USB/Bluetooth). Si tu
  lector requiere otra integración (serie, SDK propio), adapta el método
  `_leer_of()` en `main.py`.
- El identificador de operario se modela como texto libre; si dispones de
  tarjetas RFID o lector de credenciales, se puede conectar de la misma
  forma (como entrada de teclado) o integrando el SDK correspondiente.
- Los tamaños de fuente/botones (`config.py`) están pensados para pantallas
  táctiles; ajústalos según la resolución real del dispositivo.
