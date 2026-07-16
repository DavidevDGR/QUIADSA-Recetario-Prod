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

## Generar un .exe para enseñar la app en otros equipos (Windows)

Se usa **PyInstaller**. Importante: PyInstaller compila para el sistema
operativo en el que se ejecuta, así que este paso **debes hacerlo tú en tu
PC con Windows** (no se puede generar un .exe de Windows desde Linux/Mac).

1. Abre un símbolo de sistema (cmd) en la carpeta del proyecto.
2. Ejecuta:
   ```bat
   build_exe.bat
   ```
   Esto instala las dependencias (`requirements.txt`), limpia compilaciones
   anteriores y genera el ejecutable.
3. El resultado queda en `dist\AppFabricacion.exe`. Para enseñarlo en otro equipo,
   **copia solo ese fichero** — no hace falta llevar Python instalado ni el
   resto de carpetas del proyecto.

Notas:
- Se compila con `--exclude-module pyodbc` porque, mientras `MOCK_MODE =
  True`, la app no necesita esa librería. Cuando conectes con Solmicro de
  verdad (ver sección siguiente) y pases a `MOCK_MODE = False`, instala
  `pyodbc` (descomenta la línea en `requirements.txt`) y quita esa
  exclusión de `build_exe.bat` antes de recompilar.
- El primer arranque del `.exe` puede tardar unos segundos más de lo normal
  (PyInstaller descomprime la app en una carpeta temporal); los siguientes
  arranques también, ya que se usa `--onefile`. Si prefieres arranques más
  rápidos a cambio de repartir una carpeta en vez de un único fichero,
  cambia `--onefile` por `--onedir` en `build_exe.bat`.
- La base de datos (`cuba_app.db`) y la carpeta `exports\` se crean junto
  al `.exe`, no dentro de él, así que los datos persisten entre ejecuciones
  y son fáciles de localizar/borrar.

## Estructura

- `build_exe.bat` — script para compilar `AppFabricacion.exe` con PyInstaller (ver
  sección anterior).
- `requirements.txt` — dependencias del proyecto.

- `config.py` — parámetros ajustables (rutas, márgenes de tiempo y peso, conexión Solmicro).
- `db.py` — acceso a datos:
  - `SolmicroRepository`: lectura de la OF y de los "Tipos de Ruta" en Solmicro.
    **Actualmente en modo MOCK** (datos de ejemplo) para poder probar la app
    sin la conexión real. Ver la sección "Conectar con la base de datos real
    de Solmicro" más abajo para activarla paso a paso.
  - `LocalRepository`: base de datos SQLite local donde se registra toda la
    ejecución (operarios iniciales y por sección, secciones, tiempos,
    incidencias de tiempo y de peso). Es la fuente de datos para la
    exportación a Excel.
- `export.py` — genera/actualiza `exports/registro_fabricacion.xlsx` con una
  fila por sección, incluyendo: **ID autoincremental** (1, 2, 3... continúa
  donde se quedó la última fila, no se reinicia entre exportaciones), fecha
  de exportación, estado (CANCELADA/FINALIZADA), OF, **lote**
  (`tbOrdenFabricacion.Lote`, se obtiene automáticamente de Solmicro pero
  solo se exporta, nunca se muestra en la app), máquina, operarios
  iniciales, fechas de inicio/fin, tiempo total, peso tara/total/neto y su
  incidencia si la hubo, y por cada sección: instrucción, operarios activos
  al cerrarla, tiempo previsto/real, incidencia, y las pausas ocurridas
  **durante esa sección concreta** en tres columnas separadas ("Pausa
  inicio", "Pausa fin", "Pausa (min)"); si hay varias pausas en la misma
  sección, se separan con " ; " dentro de cada columna.
  - El rango de datos se formatea como una **Tabla de Excel real** (con
    encabezados, filtros y filas alternas), no solo como celdas sueltas;
    se rehace en cada exportación para cubrir las filas nuevas.
  - Los encabezados de fecha/hora usan "y" en lugar de "/" (p. ej.
    "Fecha y Hora exportación") por compatibilidad con otras aplicaciones
    que puedan interpretar "/" como separador al importar el Excel.
  - En producción, `exports/` debería ser una carpeta sincronizada con la
    nube (OneDrive/SharePoint/Google Drive) asociada a
    `fabricacion@quiadsa.com`, o bien puedes activar la subida directa vía
    API (hay un ejemplo con Microsoft Graph comentado al final del fichero).
- `main.py` — interfaz gráfica con sistema de pestañas (varias OF a la vez)
  y las 3 fases dentro de cada pestaña.
- `onscreen_keyboard.py` — teclado virtual QWERTY en pantalla, reutilizable en
  cualquier campo de texto (`attach_keyboard(entry)`) o numérico
  (`attach_keyboard(entry, numeric_only=True)`). Se llama así (y no
  `keyboard.py`) para no chocar con el paquete de PyPI `keyboard`.

## Conectar con la base de datos real de Solmicro

### 1) Qué pedir a IT / soporte de Solmicro

- Servidor (IP o nombre) y puerto, si no es el 1433 por defecto.
- Nombre de la base de datos.
- Usuario y contraseña — a ser posible un **usuario de solo lectura**,
  ya que esta aplicación únicamente necesita consultar (SELECT), nunca
  escribir en Solmicro.
- Confirmar que el PC/pantalla táctil tiene instalado el **driver ODBC de
  SQL Server** (en Windows normalmente ya viene, o se instala con el
  "ODBC Driver 17/18 for SQL Server" de Microsoft).

### 2) Configurar la conexión

Edita `config.py` y rellena `SOLMICRO_CONN_STRING` con esos datos (el
usuario/contraseña de solo lectura ya están puestos: `lectura`/`lectura`;
solo falta el servidor y la base de datos reales):

```python
SOLMICRO_CONN_STRING = (
    "DRIVER={ODBC Driver 17 for SQL Server};"
    "SERVER=IP_O_NOMBRE_SERVIDOR,1435;"   # <-- tu servidor real (puerto 1435)
    "DATABASE=NOMBRE_BASE_DATOS;"          # <-- nombre real de la BD
    "UID=lectura;"
    "PWD=lectura;"
)
```

Instala la librería necesaria:
```bash
pip install pyodbc
```

### 3) Esquema real ya mapeado

Las consultas en `config.py` ya están escritas contra el esquema real
confirmado (tablas `tbOrdenFabricacion` y `tbRuta`):

| Campo de la app | Columna real |
|---|---|
| Código OF (escaneado) | `tbOrdenFabricacion.NOrden` |
| Artículo | `tbOrdenFabricacion.IDArticulo` |
| Cantidad (kg) | `tbOrdenFabricacion.QFabricar` |
| Centro/máquina | `tbOrdenFabricacion.IDCentroGestion` |
| Tipo de ruta | `tbOrdenFabricacion.IDTipoRuta` |
| Nº de sección | `tbRuta.Secuencia` |
| Instrucción | `tbRuta.Texto` |
| Tiempo previsto (min) | `tbRuta.TiempoEjecUnit` (ya en minutos) |

Las secciones se filtran por `IDTipoRuta`. Si al probar aparecen
operaciones de otro artículo mezcladas (porque varios artículos comparten
el mismo tipo de ruta), cambia el filtro de `SQL_SECCIONES_RUTA` en
`config.py` para usar `IDArticulo + IDRuta` en su lugar, que es más
específico.

### 4) Activar el modo real

En `db.py`, cambia:
```python
MOCK_MODE = True   →   MOCK_MODE = False
```
A partir de ahí, `SolmicroRepository` se conectará de verdad a Solmicro en
cada consulta de OF, en lugar de devolver los datos de ejemplo.

## Probar la aplicación sin conexión a Solmicro (modo MOCK)

Mientras `MOCK_MODE = True` en `db.py`, puedes escribir **cualquier texto no
vacío** como código de OF (p. ej. `OF-2026-001`, `PRUEBA`, `12345`) y la
aplicación devolverá siempre la misma orden de ejemplo: artículo
`FORMULA-DEMO-A`, 500 kg, centro `CUBA-01`, con 5 secciones de prueba
(distintos tiempos previstos para poder probar el margen del 10%). No hay
que "acertar" ningún código concreto: el mock no valida nada, solo simula
que Solmicro ha respondido.

## Comportamiento clave implementado

- **Sistema de pestañas** (como un navegador): se pueden tener **varias OF
  abiertas a la vez**, cada una en su propia pestaña, con su propio estado
  (operarios, sección actual, temporizador, etc.) totalmente independiente.
  - Botón **"+ Nueva OF"** al final de la barra de pestañas para empezar
    otra orden sin perder el progreso de las demás.
  - Click en una pestaña para cambiar a ella; el temporizador de la sección
    en curso sigue corriendo "en segundo plano" aunque no se esté viendo
    esa pestaña (se recalcula a partir de la hora real de inicio en cuanto
    se vuelve a mostrar).
  - Botón **✕** en cada pestaña para cerrarla del todo; si la OF está en
    curso pide confirmación **y un motivo de cancelación** (obligatorio) antes
    de cerrar y exportar.
  - Al **finalizar o cancelar** una OF con los botones de la propia pantalla
    (Fin / Cancelar OF), la pestaña **no se cierra**: vuelve a la Fase 1 en
    blanco, lista para empezar otra OF sin perder la pestaña.
  - Si se cierra la **ventana completa de la aplicación** (botón X del
    sistema operativo) con alguna OF en curso en cualquier pestaña, se pide
    confirmación y el motivo de cada una, se cancelan y se exportan como
    CANCELADA antes de cerrar de verdad la aplicación.
- **Fase 1**: alta de uno o varios operarios (Enter o botón Añadir), con
  lista donde se pueden seleccionar y **eliminar** por si alguien se ha
  equivocado. Lectura de OF por código de barras, cuba (teclado numérico) y
  peso de tara. Todos los campos muestran el **teclado virtual** al
  pulsarlos (QWERTY para texto, numérico para cuba y pesos).
- **Fase 2**: instrucción de la sección, temporizador ascendente, botones
  Cancelar OF (arriba-izquierda), Anterior y Siguiente/Finalizar.
  - **Toda cancelación de una OF** (botón Cancelar OF, cerrar su pestaña, o
    cerrar la aplicación) **exige escribir un motivo** mediante una ventana
    con botones **Aceptar** y **Cancelar**. Si se pulsa Cancelar (o se
    cierra la ventana con el aspa), se aborda toda la acción y no se
    cancela nada. Si se acepta con un motivo, este se registra como la
    incidencia de la sección que estaba en curso en ese momento, y se
    exporta en su fila correspondiente del Excel.
  - **El botón Anterior también pide un motivo** (con las mismas opciones
    Aceptar/Cancelar) antes de volver a la sección previa; el motivo
    indicado se registra como la incidencia de la sección que se cierra al
    retroceder, y se exporta al Excel. Si se pulsa Cancelar en ese diálogo,
    no se retrocede de sección.
  - **Botón Pausa/Reanudar** (arriba a la derecha, en paralelo a Cancelar
    OF): amarillo "⏸ PAUSA" al
    pulsarlo registra la fecha/hora de inicio de la pausa y cambia a verde
    "▶ REANUDAR"; al volver a pulsarlo registra la fecha/hora de fin y la
    duración en minutos. Si se cancela o finaliza la OF con una pausa
    abierta, se cierra automáticamente para no dejar datos incompletos.
    **Si hay una pausa sin reanudar, no se puede avanzar con
    Siguiente/Finalizar ni retroceder con Anterior**: aparece un aviso
    pidiendo reanudarla primero.
    - Comportamiento configurable con `config.PAUSA_DETIENE_CONTADOR`:
      - `False` (por defecto): el temporizador de la sección **sigue
        corriendo en tiempo real** aunque esté en pausa; la pausa solo
        queda registrada para el Excel, sin afectar al cronómetro ni al
        tiempo real de la sección.
      - `True`: el temporizador **se congela** al pulsar Pausa (p. ej. en
        00:30) y continúa exactamente desde ahí al pulsar Reanudar (00:31,
        00:32...). Esto también afecta al tiempo real de la sección que se
        exporta al Excel y a la comprobación del margen de tiempo del 10%,
        que en ese caso se calculan excluyendo el tiempo en pausa.
  - **Gestión de operarios habilitada también aquí**: se pueden añadir o
    quitar operarios en cualquier sección, no solo al principio. **No se
    puede avanzar con Siguiente/Finalizar ni retroceder con Anterior si no
    queda ningún operario activo** (igual que en la Fase 1): aparece un
    aviso pidiendo añadir al menos uno. Al cerrar
    cada sección (Siguiente/Finalizar/Anterior/Cancelar) se exporta una
    "foto" de los operarios que están **activos en ese momento** en esa
    sección concreta; si alguien se quita antes de que la sección se
    cierre, ya no aparecerá en su registro.
  - Si se pulsa Siguiente con un desvío de más del **10%** (antes o
    después, `config.MARGEN_TIEMPO_PORCENTAJE`) respecto al tiempo
    previsto de la sección, se exige un motivo antes de continuar (también
    con opción de Cancelar para no avanzar).
- **Fase 3**: tiempo total, peso tara, campo de peso total con cálculo en
  vivo del **peso neto** (peso total − peso tara).
  - Si el peso neto se desvía más de un **10%** (`config.
    MARGEN_PESO_PORCENTAJE`) de la cantidad prevista en la OF, se exige un
    motivo antes de poder finalizar (con opción de Cancelar).
  - Botón Fin: registra, exporta a Excel (peso tara/total/neto, motivo de
    la incidencia de peso si la hubo, y operarios de cada sección) y
    cierra la pestaña.

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
- Los porcentajes de margen (tiempo y peso) son ajustables en `config.py`
  (`MARGEN_TIEMPO_PORCENTAJE` y `MARGEN_PESO_PORCENTAJE`).
