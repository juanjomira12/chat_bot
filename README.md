# Bot de Finanzas Personales (Telegram)

Bot: t.me/Finanzas_personales123_bot

## Archivos
- `bot.py` – código del bot
- `preguntas.json` – categorías, preguntas, palabras clave y respuestas (edítalo para agregar las tuyas)
- `.env` – token del bot (copia `.env.example` a `.env`) (`TELEGRAM_TOKEN=...`), y opcional `TELEGRAM_PROXY=...`
- `finanzas.db` – se crea sola; guarda gastos e ingresos de cada usuario

## Ejecutar
```powershell
.\venv\Scripts\python.exe bot.py
```
Detener con Ctrl+C. El bot solo responde mientras este comando está corriendo.

## Reinstalar dependencias (si borras el venv)
```powershell
python -m venv venv
.\venv\Scripts\python.exe -m pip install -r requirements.txt
```

## Uso en Telegram
Escribe /start y usa el menú de botones:
- ➖ Registrar gasto / ➕ Registrar ingreso: categoría → monto → descripción (u Omitir)
- 📊 Resumen del mes, 📋 Últimos movimientos, 🗑️ Borrar último (pide confirmación)
- 📥 Reporte Excel / 📄 Reporte Word: eliges periodo y el bot envía el archivo con gráficas
- ❓ Preguntas: temas de finanzas (también puedes escribir la pregunta)

Atajos opcionales: /gasto 25000 almuerzo, /ingreso 1500000 salario, /cancelar.
`reportes.py` genera los archivos; en `ejemplos/` hay reportes de muestra.

## Agregar preguntas
En `preguntas.json` añade un objeto a `preguntas` con `categoria` (una clave de `categorias`),
`pregunta`, `palabras_clave` y `respuesta`. Reinicia el bot para cargar los cambios.
