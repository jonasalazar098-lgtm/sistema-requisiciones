# Sistema de requisiciones v12

Cambios sobre v11:
- Mantiene la descarga del Excel generado.
- Agrega botón para descargar el PDF generado a partir del Excel final.
- El PDF se genera con LibreOffice en modo headless.
- Se agrega packages.txt con libreoffice-calc para Streamlit Community Cloud.

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

## Nota para PDF

Para que el PDF funcione en Streamlit Cloud, el repositorio debe incluir:

```txt
packages.txt
```

con este contenido:

```txt
libreoffice-calc
```
