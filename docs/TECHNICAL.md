# Documentación técnica y revisión de publicación

## Flujo y responsabilidades

1. `excel_reader.read_excel` lee la primera hoja con Pandas y openpyxl, conserva textos numéricos y valida encabezados.
2. `comparison.validate_document` valida nombres, duplicados y valores; `compare_documents` calcula estados, importes y agregados con Decimal.
3. `app.py` muestra resultados, aplica filtros sobre copias y presenta dinero con `format_money_ar`. No modifica las reglas del motor.
4. `report.create_report` recibe el resultado completo y produce bytes XLSX con openpyxl/BytesIO. No lee nuevamente los documentos ni recalcula las diferencias.

## Contrato y presentación

Encabezados exactos: Producto, Cantidad, Precio unitario. La normalización usa `casefold` y espacios reducidos; no elimina acentos. Se rechazan nombres vacíos, duplicados normalizados, booleanos, vacíos, números negativos y valores no finitos. Los textos numéricos deben usar punto decimal sin miles.

Ausentes: importe cero y cantidad/precio sin dato. Los seis estados no cambian con los filtros. Exceso y de menos mantienen signos positivo y negativo respectivamente; el neto los compensa. La precisión Decimal se ajusta a los datos y los cálculos no se redondean hasta su presentación.

La interfaz transforma la salida de `format_money` a es-AR, sin convertir a float ni cambiar el locale global: `1234.50` → `1.234,50`. Los datos originales y los valores del motor no se alteran. En Excel, una cantidad integral usa formato `0`; una cantidad fraccionaria conserva el formato decimal existente. Los códigos de formato del XLSX son independientes de la configuración regional. Los importes siguen usando dos decimales y las cifras de precisión excesiva se exportan como texto.

## Dependencias y portabilidad

Python probado: 3.12.14 en Windows. `requirements.txt` contiene Streamlit, Pandas, openpyxl y pytest con versiones exactas. `requirements-lock.txt` captura las versiones directas y transitivas instaladas, sin rutas locales ni referencias a repositorios privados. No contiene secretos ni necesita claves.

`pip check` no informó conflictos. La suite de 80 pruebas pasó. No se reinstaló desde Internet ni se verificó una instalación limpia en Linux/Cloud: la disponibilidad futura de cada versión y del intérprete deberá comprobarse antes del despliegue. El lock fija versiones, no hashes de distribución.

Los recursos se resuelven respecto de los módulos, sin rutas personales absolutas. Ejecutar los comandos y la suite desde la raíz del proyecto. No copiar ni versionar `.venv`.

## Configuración de ejecución

`.streamlit/config.toml` contiene el tema público y desactiva estadísticas de uso. La dirección `127.0.0.1` limita la ejecución actual a la máquina local. Para un futuro alojamiento, revisar esa dirección o sobrescribirla según el entorno; no se modificó ahora para evitar exponer el servidor local. El archivo de entrada será `app.py`.

Reiniciar el proceso después de cambiar módulos Python. Recargar solo el navegador puede conservar versiones de módulos anteriores en el proceso del servidor.

## Revisión de seguridad · 29/10/2026

Alcance: archivos candidatos obtenidos con `git ls-files --cached --others --exclude-standard`, incluyendo celdas y XML/metadatos de ambos Excel.

- No se detectaron credenciales, firmas de tokens, claves privadas, correos personales ni rutas de usuario en los archivos revisados.
- Los dos ejemplos contienen exclusivamente productos ficticios y sus cantidades/precios. Autor de los metadatos: openpyxl; sin último editor personal, comentarios ni hojas ocultas.
- Se comprobaron exclusiones para entornos virtuales, cachés, `.env`, secretos de Streamlit, herramientas privadas, claves, temporales y documentos locales. Las únicas excepciones XLSX son los dos ejemplos previstos.
- No hay archivos ya rastreados que eludan las nuevas exclusiones. No hay commits ni remoto configurado. No se crearon commits, conexiones remotas ni publicaciones.
- Los nombres de productos se exportan como texto, no como fórmulas Excel.

La revisión es puntual y no garantiza la ausencia absoluta de secretos. Antes de publicar cambios futuros, revisar nuevamente la lista de archivos y el diff. `.gitignore` no elimina archivos que ya estén rastreados y no debe forzarse la inclusión de documentos privados.

## Verificación y pendientes

- 80 pruebas aprobadas, incluyendo formatos de dinero y cantidades.
- Arranque local y endpoint de salud verificados; no implica que Cloud haya sido probado.
- Interfaz revisada previamente a 1280 px y 390 px. No se certificó compatibilidad con todos los teléfonos o navegadores.
- Pendientes para compartir: autorización y primer commit; decidir licencia si se quiere conceder permisos de reutilización.
- Pendientes exclusivos del despliegue: configuración del host, instalación limpia y validación funcional en Cloud.

Esta revisión prepara el código para compartirlo. No autoriza ni realiza publicación en GitHub o Streamlit Community Cloud.
