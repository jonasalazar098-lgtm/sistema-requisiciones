# Sistema de requisiciones v23

Corrección sobre v22:
- Se revierte el cambio que rompía las líneas de la tabla.
- No descombina el cuerpo completo de la requisición.
- Respeta la estructura original de la plantilla.
- Corrige la alineación usando la celda ancla real cuando hay celdas combinadas.
- Mantiene:
  - franja gris oscura debajo del último producto.
  - una sola diagonal limpia.
  - descarga en Excel y PDF.
  - filas de la tabla con líneas normales.

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
