# Sistema de requisiciones v26

Cambio principal:
- Cuando una requisición tiene más productos de los que caben en el formato, la app ya no fuerza todo en una sola hoja.
- Ahora pagina automáticamente:
  - JAG25-1
  - JAG25-2
  - etc.

Esto evita:
- que el último producto salga sin celdas o sin precio.
- que Observaciones, Facturar a, Recibe, Entregar en y Firmas se muevan.
- que se rompa la parte inferior del formato.

Mantiene:
- Excel y PDF.
- franja gris oscura debajo del último producto cuando hay espacio.
- una sola diagonal limpia en el espacio vacío.
- eliminación de hojas vacías de plantilla.

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
