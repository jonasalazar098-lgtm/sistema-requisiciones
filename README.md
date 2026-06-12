# Sistema de requisiciones v11

Cambios sobre v10:
- Se eliminó el texto técnico de la página principal.
- El título queda como: Sistema de automatización de requisiciones.
- La descripción principal queda en estilo profesional para cliente.
- Las opciones de archivo origen ahora aparecen solo como Excel y PDF.
- Mantiene FORMATOS.xlsx como banco de plantillas interno.
- El Excel final conserva solo la hoja generada.

## Ejecutar

```powershell
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python -m streamlit run app.py
```
