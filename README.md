# Sistema de requisiciones v27

Cambio sobre v26:
- Cuando la requisición se divide en varias hojas, el bloque inferior solo aparece en la última hoja.
- En hojas intermedias se ocultan:
  - Observaciones.
  - Facturar a.
  - Recibe.
  - Entregar en.
  - Firmas.
  - Cargar a / ID Familia / Cargo / Área / Partida.

Mantiene:
- Paginación automática cuando hay muchas partidas.
- Excel y PDF.
- Franja gris oscura debajo del último producto.
- Una sola diagonal limpia cuando hay espacio vacío.
- Eliminación de hojas vacías de plantilla.

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
