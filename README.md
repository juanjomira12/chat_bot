# Bot de Finanzas Personales (Telegram)

Bot: t.me/Finanzas_personales123_bot

## Archivos
- `bot.py` – código del bot
- `preguntas.json` – categorías, preguntas, palabras clave y respuestas (edítalo para agregar las tuyas)
- `reportes.py` – genera los reportes en Excel y Word con gráficas
- `.env` – configuración local; copia `.env.example` a `.env` y pon tu token (`TELEGRAM_TOKEN=...`)
- `finanzas.db` – se crea sola; guarda gastos e ingresos de cada usuario (ruta cambiable con `DB_PATH`)
- `Procfile`, `.python-version` – para desplegar en un servidor

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

## Desplegar en Railway (24/7, sin depender de tu PC)
1. Sube este repo a GitHub.
2. En railway.com: **New Project → Deploy from GitHub repo** → elige `chatbot`.
3. En el servicio, pestaña **Variables**, agrega:
   - `TELEGRAM_TOKEN` = el token de BotFather
   - `DB_PATH` = `/data/finanzas.db`
4. Clic derecho en el servicio → **Attach Volume** con ruta de montaje `/data`
   (sin volumen, los movimientos se borran en cada despliegue).
5. Railway lee el `Procfile` y ejecuta `python bot.py`. En **Deploy Logs** debe aparecer
   `✅ Bot MIRA conectado como @...`.

Importante: el bot solo puede correr en **un lugar a la vez**. Cuando esté en el servidor,
no lo ejecutes también en tu PC (Telegram da error *Conflict*).

Otros servicios (Render, Fly.io, un VPS) sirven igual: instalar `requirements.txt`, definir
las variables de entorno, ejecutar `python bot.py` como proceso de fondo (worker) y usar
un disco persistente para `DB_PATH`.

## Agregar preguntas
En `preguntas.json` añade un objeto a `preguntas` con `categoria` (una clave de `categorias`),
`pregunta`, `palabras_clave` y `respuesta`. Reinicia el bot para cargar los cambios.
