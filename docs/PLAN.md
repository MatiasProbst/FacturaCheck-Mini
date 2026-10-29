# Plan de desarrollo

Cada etapa se iniciará cuando el usuario la solicite. Al finalizar se explicarán los cambios, las verificaciones realizadas y lo que quede pendiente. No avanzar automáticamente a la siguiente etapa.

## Publicación del código en GitHub · 29/10/2026

El usuario autoriza el primer commit `Initial release - FacturaCheck Mini` y la
publicación de `main` en `MatiasProbst/FacturaCheck-Mini`. Se revisan los 19 archivos,
las exclusiones y el historial, y se ejecutan las 80 pruebas antes de publicar.
No se sobrescribe contenido remoto ni se utiliza push forzado. No se agrega licencia.
El despliegue en Streamlit Community Cloud continúa pendiente de autorización.

## Preparación local para GitHub · 29/10/2026

Revisión de estructura, exclusiones, dependencias, ejemplos y datos sensibles.
README orientado a visitantes, documentación técnica y lock de versiones del entorno.
Presentación monetaria es-AR en interfaz y cantidades enteras sin decimales en Excel.
El motor de comparación permanece sin cambios. No se crean commits, remotos ni
publicaciones; se espera autorización. Cloud continúa pendiente de validación y despliegue.

## Etapa 5: interfaz

Tema claro y componentes nativos, secciones de carga/resumen/detalle, tarjetas,
ejemplos descargables y datos originales desplegables. Filtros por estado y solo
diferencias sobre la tabla visible, sin afectar el informe completo. Sección Acerca
del proyecto con validaciones y límites. 78 pruebas aprobadas al completar estos
cambios; revisión visual en escritorio y viewport móvil de 390 px. No se agregaron
dependencias ni se modificó el motor. La preparación posterior añade dos regresiones
de formato, para un total de 80 pruebas.

## Antecedente: etapa 4 solicitada

Se incorpora el informe descargable con hojas Resumen, Comparación y Observaciones.
`report.py` usa openpyxl y BytesIO, sin archivos temporales ni cambios en los
originales. Las observaciones se generan por reglas. El motor calcula los agregados
Decimal una sola vez y los comparte con interfaz y reporte: exceso (positivo),
de menos (negativo), neto y totales solicitado/facturado. Se agregan dos indicadores
y el botón de descarga, con pruebas de contenido, formatos e integración.
No se agregan dependencias. Publicación y otras etapas requieren nueva instrucción.

## Antecedente: etapa 3 solicitada

Se incorpora la comparación por Producto, validación de contenido, seis estados,
tabla e indicadores. La lógica queda en `comparison.py`, separada de Streamlit.
Se usa Decimal sin redondeos intermedios; la presentación redondea a dos decimales
con ROUND_HALF_UP. Ausentes: importe cero y cantidad/precio sin dato. Los nombres
se normalizan con casefold y reducción de espacios. Se rechazan duplicados,
productos vacíos y cantidades/precios no finitos, negativos o no numéricos.
El README documenta el contrato completo. Se mantienen las dependencias existentes.
En esa entrega quedaron pendientes reportes y publicación; los reportes se
incorporan en la etapa 4 descrita arriba y la publicación sigue pendiente.

## Antecedente: etapa 2 solicitada

El usuario autorizó adelantar una primera interfaz Streamlit junto con la lectura
y validación estructural. Esta entrega incluye dos cargas independientes, tablas,
mensajes de error, ejemplos ficticios y pruebas. Se lee la primera hoja de cada
`.xlsx`, con encabezados exactos `Producto`, `Cantidad` y `Precio unitario` en la
primera fila; se permiten columnas adicionales. Los encabezados sin filas se
aceptan con una advertencia.

En esa entrega las reglas de contenido, la comparación y los reportes quedaron
pendientes; contenido y comparación se incorporan en la etapa 3 descrita arriba.
Las etapas originales que siguen
se conservan como referencia; la interfaz completa de la etapa 6 sigue pendiente.
No avanzar sin una nueva instrucción.

## 1. Base del proyecto

- Crear instrucciones permanentes y descripción inicial.
- Registrar dependencias previstas y exclusiones de Git.
- Inicializar Git local si no existe.
- Verificar los archivos generados y el estado del repositorio.

Criterio de finalización: documentación y configuración inicial presentes, sin código de aplicación.

## 2. Entorno y contrato de datos

- Preparar un entorno virtual en Windows e instalar dependencias con autorización para el acceso externo necesario.
- Seleccionar y registrar versiones probadas de Python y bibliotecas.
- Definir columnas, identificador de producto, campos obligatorios y formatos aceptados para ambos Excel.
- Acordar reglas para duplicados, valores vacíos, cantidades, moneda, precisión, redondeo y tolerancias monetarias.
- Definir el alcance de impuestos y descuentos antes de implementar reglas relacionadas.
- Acordar una estructura modular mínima y preparar pytest.

Criterio de finalización: entorno verificado y contrato documentado con datos ficticios.

## 3. Lectura y validación

- Leer archivos `.xlsx` y validar estructura y contenido según el contrato.
- Informar errores de forma comprensible.
- Incorporar archivos de ejemplo ficticios y pruebas para archivos válidos e inválidos.

Criterio de finalización: lectura y validaciones cubiertas por pruebas ejecutadas.

## 4. Comparación

- Relacionar productos según el identificador acordado.
- Detectar productos faltantes o adicionales y diferencias de cantidades y precios.
- Aplicar una representación monetaria exacta y las reglas de redondeo acordadas.
- Probar coincidencias, diferencias y casos límite.

Criterio de finalización: reglas de comparación implementadas y pruebas relevantes aprobadas.

## 5. Reporte descargable

- Definir una tabla clara de resultados y un reporte Excel.
- Verificar mediante pruebas el contenido del archivo generado y su lectura posterior.

Criterio de finalización: reporte generado y contenido comprobado con ejemplos ficticios.

## 6. Interfaz Streamlit

- Incorporar carga de los dos archivos, mensajes de validación, tabla de resultados y descarga.
- Reutilizar los módulos de lectura, validación, comparación y reporte.
- Comprobar manualmente el flujo completo local en Windows y ejecutar las pruebas disponibles.

Criterio de finalización: flujo probado y limitaciones documentadas.

## 7. Preparación para publicación

- Completar README con instalación, ejecución, pruebas, contrato de datos y ejemplos.
- Revisar dependencias reproducibles, ausencia de información sensible y compatibilidad con Streamlit Community Cloud.
- Ejecutar la suite de pruebas y documentar resultados.
- Publicar en GitHub y Streamlit Community Cloud solo con autorización explícita.

Criterio de finalización: proyecto reproducible y listo para compartir; publicación verificada únicamente si fue autorizada y realizada.

