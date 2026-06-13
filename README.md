# Sistema de requisiciones v20

Corrección sobre v19:
- Corrige que, cuando hay muchas partidas, los últimos productos salgan centrados o desacomodados.
- Las filas extra ahora copian el formato de una fila normal de producto.
- Fuerza alineación estable:
  - cantidades, unidades y códigos centrados.
  - descripción alineada a la izquierda.
  - precios e importes a la derecha.
- Mantiene franja gris oscura debajo del último producto.
- Mantiene una sola diagonal limpia en el espacio vacío.
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
