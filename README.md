# PRE-S&OP Dashboard

Aplicación Streamlit lista para subir a GitHub y desplegar en Streamlit Community Cloud.

## Archivos
- `app.py`: aplicación principal
- `requirements.txt`: dependencias

## Ejecución local
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Despliegue en Streamlit Cloud
1. Sube esta carpeta a un repositorio en GitHub.
2. En Streamlit Community Cloud, crea una nueva app.
3. Selecciona el repositorio y como archivo principal usa `app.py`.
4. Deploy.

## Notas
- La app espera archivos Excel `.xlsx` para Demanda, Abastecimiento y opcionalmente Minuta.
- No requiere cambios adicionales para Streamlit Cloud.
