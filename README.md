# Asociación de Fútbol Nacimiento

## Estructura requerida

```text
proyecto/
├── app.py
├── logoanfa.png           # también acepta jpg, jpeg, webp o gif
├── requirements.txt
└── data/
    └── Planilla_Maestra_Campeonato_AF_Nacimiento_2026.xlsx
```

La aplicación detecta automáticamente el primer archivo `.xlsx` dentro de `data` si el nombre cambia.

## Actualización de datos

1. Modifica y guarda la planilla Excel.
2. Reemplaza el archivo de la carpeta `data` en GitHub.
3. En la página, presiona **Actualizar datos**.

## Ejecución local

```bash
pip install -r requirements.txt
streamlit run app.py
```
