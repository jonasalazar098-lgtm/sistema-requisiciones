# Sistema de requisiciones v13

Cambios sobre v12:
- Mantiene la descarga de Excel.
- Mantiene la descarga de PDF.
- Agrega cierre visual después del último artículo:
  - línea gris al final del último artículo.
  - diagonal sobre el espacio vacío restante.
- El cierre se aplica en el Excel y también sale igual en el PDF, porque el PDF se genera desde el Excel final.
- Mantiene packages.txt con libreoffice-calc para Streamlit Community Cloud.

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
