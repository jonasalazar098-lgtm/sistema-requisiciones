# Sistema de requisiciones v22

Corrección sobre v21:
- Revierte el cambio que rompía las líneas de la tabla.
- Corrige los últimos productos cuando hay muchas partidas sin desarmar toda la plantilla.
- Rehace únicamente el área combinada de DESCRIPCIÓN por fila.
- Mantiene:
  - descripciones alineadas a la izquierda.
  - cantidades/unidad/código centrados.
  - precios e importes a la derecha.
  - franja gris oscura debajo del último producto.
  - una sola diagonal limpia en el espacio vacío.
  - descarga en Excel y PDF.

## Archivos requeridos

- app.py
- requirements.txt
- packages.txt
- README.md
- FORMATOS.xlsx
- .gitignore
- .streamlit/config.toml

## Ejecutar local

```powershell
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python -m streamlit run app.py
```

## Nota

Después de subir a GitHub, reiniciar la app en Streamlit Cloud:
Manage app > Reboot app
