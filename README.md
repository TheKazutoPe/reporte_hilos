# Validación de Enlaces de Fibra Óptica (Flask + Supabase)

## Qué trae esta versión (lo que pediste)
- ✅ Lista desplegable de enlaces (desde `data/enlaces_programados.json`, generado desde tu Excel)
- ✅ Origen A y Origen B como desplegables (autocompletado con `<datalist>`)
- ✅ Sugerencia automática de Origen A/B según el enlace seleccionado (split por "-")
- ✅ Opción **Editar** (cabecera + A + B + hilos) desde dashboard y reporte
- ✅ Vista tipo Excel en el reporte
- ✅ Exportación a Excel: `/enlace/<id>/export.xlsx`

## 1) Crear tablas en Supabase
En Supabase → SQL Editor → ejecuta `init.sql`.

## 2) Configurar variables de entorno
Copia `.env.example` a `.env` y completa:
- `SUPABASE_URL`
- `SUPABASE_SERVICE_ROLE_KEY` (service_role)

## 3) Instalar y correr
```bash
python -m venv .venv
# Windows:
.venv\Scripts\activate

pip install -r requirements.txt
python app.py
```

Abre: http://localhost:5000

## 4) Actualizar lista de enlaces
Reemplaza `data/enlaces_programados.json` por uno nuevo (o vuelve a generar desde tu Excel).
