# Sistema de requisiciones v17

Corrección sobre v16:
- Agrega una franja gris visible justo debajo del último producto.
- Mantiene la diagonal en el espacio vacío restante.
- El cierre visual se aplica en el Excel y también aparece en el PDF porque el PDF se genera desde el Excel final.

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
