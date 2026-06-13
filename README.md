# Sistema de requisiciones v16

Corrección sobre v15:
- Corrige el error:
  AttributeError: 'MergedCell' object has no attribute 'column_letter'
- Mantiene el cierre visual:
  - línea gris debajo del último producto.
  - diagonal en el espacio vacío.
- Mantiene descarga en Excel y PDF.
- El PDF se genera desde el Excel final.

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
