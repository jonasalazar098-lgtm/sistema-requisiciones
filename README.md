# Sistema de requisiciones v21

Corrección sobre v20:
- Corrige el problema donde, al haber muchas partidas, los últimos productos salían centrados o desacomodados.
- Antes algunas filas de la parte baja de la tabla conservaban celdas combinadas o estilos raros de la plantilla.
- Ahora se normaliza toda la zona de partidas antes de escribir:
  - se descombinan celdas del cuerpo de la tabla.
  - se limpia todo el cuerpo de partidas.
  - se reaplica formato estable.
  - descripciones quedan alineadas a la izquierda.
  - cantidades, unidades y códigos quedan centrados.
  - precios e importes quedan a la derecha.
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
