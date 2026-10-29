# FacturaCheck Mini

**Compará órdenes de compra y facturas, detectá inconsistencias y generá informes automáticamente.**

Aplicación de portfolio desarrollada con Python y Streamlit. Ayuda a revisar diferencias de productos, cantidades y precios entre dos archivos Excel estructurados, con reglas explícitas y cálculos monetarios con `Decimal`.

## El problema

La revisión manual de una factura contra una orden de compra puede dejar pasar precios distintos, cantidades incorrectas o productos ausentes. FacturaCheck Mini reúne esas diferencias en una tabla y un informe Excel para facilitar su revisión.

## Demo online

**Próximamente.** Todavía no existe una demo pública. Este espacio se actualizará cuando se autorice y verifique el despliegue.

## Funcionalidades

- Carga independiente de orden de compra y factura en `.xlsx`.
- Validación de columnas, productos duplicados y valores inválidos.
- Seis estados: correcto, diferencia de precio, diferencia de cantidad, ambas diferencias, solo en orden y solo en factura.
- Tarjetas con productos analizados, correctos, con diferencias, exceso, importe de menos y diferencia neta.
- Importes en formato español de Argentina: `1.234,56`.
- Filtro por estado y opción para mostrar solo productos con diferencias.
- Informe descargable con hojas **Resumen**, **Comparación** y **Observaciones**.
- Descarga de dos ejemplos ficticios desde la propia interfaz.
- Diseño con componentes nativos que se apilan en pantallas pequeñas; la tabla permite desplazamiento horizontal.

Los filtros afectan únicamente la tabla visible. Las tarjetas y el informe siempre incluyen la comparación completa. No se utiliza IA, bases de datos ni APIs externas.

## Tecnologías

| Tecnología | Uso |
| --- | --- |
| Python 3.12 | Lenguaje y biblioteca estándar `decimal` |
| Streamlit | Interfaz web |
| Pandas | Lectura y procesamiento de tablas |
| openpyxl | Lectura, generación y formato Excel |
| pytest | Pruebas automatizadas |
| Git | Control de versiones |

`requirements.txt` fija las dependencias directas, incluido pytest. `requirements-lock.txt` registra también las dependencias transitivas del entorno probado. No hacen falta claves de API ni servicios pagos.

## Ejecutar localmente

Requisitos: Python 3.12 y una terminal. Descargar el código y abrir una terminal en la carpeta del proyecto. La instalación de dependencias requiere acceso al índice de paquetes de Python.

### Windows · PowerShell

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements-lock.txt
.\.venv\Scripts\python.exe -m streamlit run app.py
```

No es necesario activar el entorno ni cambiar la política de PowerShell. Si Python se instaló con el lanzador de Windows, se puede usar `py -3.12 -m venv .venv` para el primer paso.

### macOS / Linux

```bash
python3.12 -m venv .venv
.venv/bin/python -m pip install -r requirements-lock.txt
.venv/bin/python -m streamlit run app.py
```

Abrir [http://127.0.0.1:8501](http://127.0.0.1:8501). Detener con `Ctrl+C`. Después de actualizar módulos Python, detener y volver a iniciar el servidor para evitar módulos anteriores en memoria.

La ejecución y las pruebas se verificaron en Windows con Python 3.12.14. Los comandos para otros sistemas no se ejecutaron en esta revisión. El lock reproduce las versiones del entorno existente; aún no se verificó una instalación limpia en otra máquina.

## Ejemplo de uso

1. Descargar los dos ejemplos desde la interfaz o usar los de `examples/`.
2. Cargar `orden_compra.xlsx` en **Orden de compra** y `factura_proveedor.xlsx` en **Factura del proveedor**.
3. Revisar el resumen y filtrar el detalle según el estado que interese.
4. Seleccionar **Descargar informe Excel**.

Resultado esperado con los ejemplos incluidos:

| Indicador | Resultado |
| --- | ---: |
| Productos analizados | 4 |
| Correctos | 0 |
| Con diferencias | 4 |
| Facturado en exceso | +1.925,00 |
| Facturado de menos | −5.150,50 |
| Diferencia neta | −3.225,50 |
| Total solicitado | 23.960,00 |
| Total facturado | 20.734,50 |

Todos los productos y valores de los ejemplos son ficticios.

## Formato de entrada

Se lee la **primera hoja** del archivo `.xlsx`. La primera fila debe tener estos encabezados exactos, en cualquier orden:

| Producto | Cantidad | Precio unitario |
| --- | ---: | ---: |
| Cuaderno ficticio A | 10 | 1250.50 |

Se permiten columnas adicionales. Producto debe ser texto no vacío. Las cantidades y precios deben ser numéricos, finitos y no negativos; se admiten cero y cantidades fraccionarias. Para números almacenados como texto se requiere punto decimal y ningún separador de miles. **El formato argentino de salida no cambia el contrato de entrada.**

Los nombres se comparan sin distinguir mayúsculas y reduciendo espacios repetidos. Se rechazan duplicados después de normalizar. Un error en cualquiera de los archivos bloquea la comparación. Una hoja con encabezados y sin filas se acepta con una advertencia.

## Criterios monetarios

- Importe = cantidad × precio unitario.
- Diferencia económica = facturado − solicitado.
- Exceso: suma de diferencias positivas. De menos: suma de diferencias negativas, con signo menos. Neto: suma de ambas.
- Un producto ausente tiene importe cero en el documento donde no aparece; cantidad y precio se muestran sin dato.
- Se utiliza `Decimal` sin redondeos intermedios. La presentación usa dos decimales y `ROUND_HALF_UP`.
- Los totales se calculan antes del redondeo y pueden diferir de sumar importes ya mostrados. Una diferencia menor a un centavo puede mostrarse como `0,00` y seguir siendo una diferencia.

El informe reutiliza los resultados del motor. Sus cantidades enteras se muestran sin separador decimal; las fracciones se conservan. Excel adapta los separadores numéricos a la configuración regional de quien lo abre.

## Estructura

```text
FacturaCheck-Mini/
├── app.py                   # Interfaz, filtros y presentación
├── excel_reader.py          # Lectura y validación de encabezados
├── comparison.py            # Validación de contenido y comparación Decimal
├── report.py                # Informe Excel en memoria y observaciones por reglas
├── requirements.txt         # Dependencias directas
├── requirements-lock.txt    # Versiones completas del entorno probado
├── .streamlit/config.toml   # Tema y configuración local
├── .gitignore
├── AGENTS.md                # Instrucciones de colaboración
├── examples/                # Dos Excel ficticios
├── scripts/create_examples.py
├── tests/                   # Lectura, comparación, interfaz e informe
└── docs/
    ├── PLAN.md
    └── TECHNICAL.md          # Contrato técnico y revisión previa a publicación
```

## Pruebas

Windows:

```powershell
.\.venv\Scripts\python.exe -m pytest -q
.\.venv\Scripts\python.exe -m pip check
```

macOS / Linux: reemplazar el ejecutable por `.venv/bin/python`.

Última verificación, **29/10/2026: 80 pruebas aprobadas**. Incluye estados, normalización, precisión, errores de entrada, filtros, descarga, formato argentino e integridad del informe. AppTest simula los selectores de archivos y ejecuta la lectura, comparación y visualización reales. La etapa de interfaz también se revisó en navegador de escritorio y en una vista de 390 px, sin sustituir una prueba en dispositivos físicos.

## Limitaciones y privacidad

- No interpreta facturas PDF, imágenes ni formatos de Excel arbitrarios.
- Supone la misma moneda; no calcula impuestos, descuentos ni conversiones.
- No consolida productos duplicados ni permite seleccionar otras hojas.
- Excel limita la precisión de celdas numéricas; no se recuperan dígitos perdidos en el origen. En la exportación, cifras que exceden la precisión admitida se preservan como texto decimal.
- No hay autenticación, persistencia ni historial. Los archivos se procesan en memoria del servidor y no se modifican los originales. En un futuro despliegue, los archivos se enviarían al servidor de alojamiento: no cargar información confidencial en una demo pública.
- No se realizaron pruebas de carga masiva ni una auditoría de seguridad independiente.

`.gitignore` excluye entornos, cachés, secretos, configuraciones privadas y documentos locales; solo permite los dos Excel ficticios previstos. Revisar siempre los archivos antes de incorporarlos a Git.

## Estado de publicación

Repositorio del proyecto: [MatiasProbst/FacturaCheck-Mini](https://github.com/MatiasProbst/FacturaCheck-Mini). La publicación del código está autorizada; todavía no existe una demo online. Para Streamlit Community Cloud, el punto de entrada previsto es `app.py` y están declaradas las bibliotecas necesarias. Antes del despliegue se deberá verificar la instalación en ese entorno y ajustar la dirección local de escucha; **Cloud no fue probado**.

No se ha elegido una licencia de reutilización. Es una decisión pendiente del autor antes de presentar el proyecto como software de código abierto.

Más detalles: [documentación técnica](docs/TECHNICAL.md) y [plan de trabajo](docs/PLAN.md).
