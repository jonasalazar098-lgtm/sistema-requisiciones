# Sistema de requisiciones v24

Corrección sobre v23:
- Se regresa a la base estable de la tabla.
- Corrige filas extra cuando hay muchas partidas copiando también las celdas combinadas de una fila normal.
- Evita que los últimos productos salgan desacomodados o sin estructura.
- Respeta las líneas y celdas de la plantilla.
- Mantiene:
  - franja gris oscura debajo del último producto.
  - una sola diagonal limpia.
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
