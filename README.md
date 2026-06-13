# Sistema de requisiciones v19

Cambios sobre v18:
- La franja gris debajo del último producto ahora es más oscura.
- La diagonal también queda en gris más oscuro.
- Mantiene una sola diagonal limpia, sin diagonales pequeñas entre celdas.
- Mantiene descarga en Excel y PDF.

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
