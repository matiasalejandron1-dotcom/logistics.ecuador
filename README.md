# 🌷 Control de Importaciones y Entregas

App en Streamlit para filtrar y organizar tus importaciones y entregas semanales a partir de tu archivo de ventas (Excel).

## ¿Qué hace?

- Subes tu archivo Excel actualizado (mismo formato de `VENTAS_2026.xlsx`, con columnas `CUSTOMER`, `VARIETY`, `PLANTS`, `IMPORT WEEK`, `DELIVERY WEEK`, etc.)
- La app clasifica automáticamente cada línea:
  - **IMPORTA Y ENTREGA (directo)**: la semana de importación es igual a la de entrega.
  - **IMPORTA Y ENRAIZA**: se importa antes; la planta se enraíza y se entrega en una semana posterior. Se calcula cuántas semanas de enraizamiento toma.
  - **REVISAR**: la entrega aparece antes que la importación (probable error de digitación en el Excel original), o la semana no se pudo interpretar.
- Puedes filtrar por cliente, variedad (búsqueda de texto), origen (`CUTTINGS FROM`), tipo de proceso, y ver todo o una semana específica.
- Al elegir una semana específica, la app te muestra en pestañas separadas:
  - 📦 Qué se importa esa semana
  - 🚚 Qué se entrega esa semana
  - 🌱 Qué lotes están enraizando (ya importados, pendientes de entrega)
- Puedes descargar el resultado filtrado en un nuevo Excel.

## Uso local

```bash
pip install -r requirements.txt
streamlit run app.py
```

Se abrirá en `http://localhost:8501`. Ahí subes tu archivo Excel cada semana (botón "Sube tu archivo Excel").

## Desplegar en Streamlit Community Cloud (gratis, conectado a GitHub)

1. Crea un repositorio nuevo en GitHub (por ejemplo `control-importaciones`).
2. Sube estos tres archivos al repositorio: `app.py`, `requirements.txt`, `README.md`.
   - Puedes arrastrarlos directamente en la interfaz web de GitHub ("Add file" → "Upload files"), o usar Git desde tu computador.
3. Entra a [share.streamlit.io](https://share.streamlit.io/) con tu cuenta de GitHub.
4. Clic en "New app", selecciona el repositorio, la rama (`main`) y el archivo principal (`app.py`).
5. Clic en "Deploy". En un par de minutos tendrás una URL pública (algo como `https://control-importaciones.streamlit.app`) que puedes abrir desde el celular o compartir con tu equipo.
6. Cada semana, entras a esa URL y subes tu Excel actualizado — no necesitas volver a desplegar nada; la app siempre usa el archivo que subas en ese momento.

## Actualizar la app después de cambios en el código

Si más adelante quieres que yo (o alguien más) ajuste el código, basta con subir la nueva versión de `app.py` al mismo repositorio de GitHub — Streamlit Cloud redespliega automáticamente.

## Notas sobre el archivo de entrada

- La app tolera espacios extra en encabezados y pequeños errores de formato en las semanas (por ejemplo `23/027` se interpreta como `23/2027`).
- Si faltan columnas clave (`CUSTOMER`, `VARIETY`, `PLANTS`, `IMPORT WEEK`, `DELIVERY WEEK`), la app te avisará con un mensaje claro en vez de fallar.
- Las filas sin semana válida o con entrega antes de importación se señalan aparte para que las revises en el Excel original.
