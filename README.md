# Sistema de requisiciones v15

Cambios sobre v14:
- Refuerza el cierre visual.
- Agrega línea gris fuerte después del último artículo.
- Agrega diagonal como shape/dibujo interno del XLSX.
- Agrega diagonales por celdas como respaldo para que se vean también en PDF.
- El PDF se genera desde el mismo Excel final.

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

Después de subir a GitHub, reinicia la app en Streamlit Cloud para forzar que use el app.py nuevo.
