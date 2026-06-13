# Sistema de requisiciones v14

Cambios sobre v13:
- Corrige el cierre visual para que sí aparezca en Excel y PDF.
- La diagonal ahora se inserta como imagen transparente sobre el espacio vacío.
- Mantiene la línea gris después del último artículo.
- El PDF se genera desde el mismo Excel final, por eso ambos deben verse iguales.
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

## Nota

Si una requisición llena casi todas las filas disponibles, puede quedar solo la línea gris y poca o ninguna diagonal porque ya no queda espacio vacío que cerrar. En requisiciones con espacio libre, aparece la diagonal sobre ese espacio.
