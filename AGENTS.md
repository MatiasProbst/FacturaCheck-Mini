# Instrucciones permanentes de FacturaCheck Mini

## Objetivo

Construir una aplicación sencilla para comparar órdenes de compra y facturas de proveedores en Excel, detectando diferencias de productos, precios y cantidades. El proyecto sirve para aprender Python y desarrollo con Codex, y posteriormente publicarse en GitHub y Streamlit Community Cloud como muestra profesional.

## Forma de trabajo

- Explicar brevemente qué se hará y por qué; el usuario está aprendiendo.
- Trabajar en etapas pequeñas. No desarrollar toda la aplicación en una sola ejecución.
- Implementar únicamente el alcance autorizado y esperar la siguiente instrucción al terminar cada etapa solicitada.
- Priorizar código modular, sencillo y mantenible, sin abstracciones innecesarias.
- Solicitar aprobación antes de realizar cambios importantes en la arquitectura.
- Antes de incorporar una tecnología nueva, explicar su propósito y obtener autorización si queda fuera de las tecnologías previstas.
- Ejecutar las pruebas disponibles relevantes y comunicar sus resultados y limitaciones. Si no hay pruebas, indicarlo.
- No afirmar que una funcionalidad funciona si no fue probada.
- No sobrescribir trabajo existente ni crear commits o publicar contenido sin autorización.

## Tecnologías y restricciones

- Usar Python, Streamlit, Pandas, openpyxl, pytest y Git.
- Mantener compatibilidad con ejecución local en Windows y futura publicación en Streamlit Community Cloud.
- Evitar dependencias adicionales que no sean necesarias.
- No requerir servicios pagos ni claves de API.
- No conectarse a servicios externos sin autorización, incluyendo instalación de dependencias y publicación.
- Usar exclusivamente datos ficticios en ejemplos y pruebas; nunca datos confidenciales ni información real de clientes.
- Evitar errores de precisión monetaria: usar `decimal.Decimal` o importes enteros en unidades mínimas, con reglas explícitas de redondeo. No basar las comparaciones monetarias en igualdad de números `float`.
- Separar progresivamente lectura, validación, comparación, reporte e interfaz conforme se aprueben las etapas.
- Cubrir con pruebas automatizadas las reglas de validación y comparación, incluidos casos inválidos y límites monetarios.

## Alcance inicial

La primera etapa crea únicamente `AGENTS.md`, `README.md`, `docs/PLAN.md`, `requirements.txt`, `.gitignore` y el repositorio Git local. No implementar la interfaz ni las funcionalidades de comparación en esta etapa.

