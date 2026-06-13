# Sistema de requisiciones v25

Corrección sobre v24:
- Corrige que, cuando hay muchas partidas, la parte inferior de la requisición se desacomode.
- Si hay más productos que filas disponibles, la app inserta filas pero comprime ligeramente la altura del cuerpo de productos.
- Esto evita que se pierdan o se deformen secciones inferiores como:
  - Observaciones.
  - Facturar a.
  - Recibe.
  - Entregar en.
  - Firmas.
- Mantiene:
  - líneas normales de la tabla.
  - descripciones alineadas correctamente.
  - franja gris oscura debajo del último producto cuando hay espacio.
  - una sola diagonal limpia cuando queda espacio vacío.
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
