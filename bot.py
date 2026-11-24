"""MIRA – Bot de Telegram de finanzas personales.

- Menú con botones para registrar gastos e ingresos (sin escribir comandos).
- Resumen del mes y reportes descargables en Excel y Word con gráficas.
- Responde preguntas frecuentes cargadas desde preguntas.json.
"""

import asyncio
import difflib
import html
import json
import logging
import os
import re
import sqlite3
import time
import unicodedata
import warnings
from datetime import datetime
from pathlib import Path

from dotenv import load_dotenv
from telegram import (
    InlineKeyboardButton,
    InlineKeyboardMarkup,
    ReplyKeyboardMarkup,
    Update,
)
from telegram.constants import ChatAction, ParseMode
from telegram.error import NetworkError, TimedOut
from telegram.ext import (
    Application,
    CallbackQueryHandler,
    CommandHandler,
    ContextTypes,
    ConversationHandler,
    MessageHandler,
    filters,
)
from telegram.warnings import PTBUserWarning

import reportes

BASE = Path(__file__).resolve().parent
load_dotenv(BASE / ".env")
TOKEN = os.getenv("TELEGRAM_TOKEN")
DB_PATH = BASE / "finanzas.db"
LOGO = BASE / "logo.png"

logging.basicConfig(format="%(asctime)s %(levelname)s %(message)s", level=logging.INFO)
logging.getLogger("httpx").setLevel(logging.WARNING)
warnings.filterwarnings("ignore", category=PTBUserWarning)
log = logging.getLogger("mira_bot")

# ------------------------------------------------------------ menú principal

B_GASTO, B_INGRESO = "➖ Registrar gasto", "➕ Registrar ingreso"
B_RESUMEN, B_ULTIMOS = "📊 Resumen del mes", "📋 Últimos movimientos"
B_EXCEL, B_WORD = "📥 Reporte Excel", "📄 Reporte Word"
B_FAQ, B_BORRAR = "❓ Preguntas", "🗑️ Borrar último"
BOTONES_MENU = [B_GASTO, B_INGRESO, B_RESUMEN, B_ULTIMOS, B_EXCEL, B_WORD, B_FAQ, B_BORRAR]

MENU = ReplyKeyboardMarkup(
    [[B_GASTO, B_INGRESO], [B_RESUMEN, B_ULTIMOS], [B_EXCEL, B_WORD], [B_FAQ, B_BORRAR]],
    resize_keyboard=True,
    is_persistent=True,
    input_field_placeholder="Elige una opción o escribe tu pregunta",
)

CATEGORIAS_MOV = {
    "gasto": {
        "Comida": "🍔", "Transporte": "🚌", "Vivienda": "🏠", "Servicios": "💡",
        "Salud": "🏥", "Entretenimiento": "🎉", "Educación": "📚", "Compras": "🛍️",
        "Deudas": "💳", "Otros": "📦",
    },
    "ingreso": {
        "Salario": "💼", "Negocio": "🏪", "Freelance": "💻", "Regalo": "🎁",
        "Inversiones": "📈", "Otros": "📦",
    },
}

# Estados de la conversación de registro
CATEGORIA, MONTO, DESCRIPCION = range(3)

# ---------------------------------------------------------------- datos (FAQ)

with open(BASE / "preguntas.json", encoding="utf-8") as f:
    DATOS = json.load(f)
CATEGORIAS_FAQ: dict[str, str] = DATOS["categorias"]
PREGUNTAS: list[dict] = DATOS["preguntas"]


def normalizar(texto: str) -> str:
    """Minúsculas y sin tildes, para comparar textos."""
    texto = unicodedata.normalize("NFD", texto.lower())
    return "".join(c for c in texto if unicodedata.category(c) != "Mn")


def buscar_pregunta(texto: str) -> dict | None:
    """Devuelve la pregunta que mejor coincide con el texto del usuario."""
    consulta = normalizar(texto)
    palabras = set(re.findall(r"[\w/]+", consulta))
    mejor, mejor_puntaje = None, 0.0
    for p in PREGUNTAS:
        claves = {normalizar(k) for k in p.get("palabras_clave", [])}
        coincidencias = sum(1 for k in claves if k in palabras or (" " in k and k in consulta))
        similitud = difflib.SequenceMatcher(None, consulta, normalizar(p["pregunta"])).ratio()
        puntaje = coincidencias + similitud * 2
        if puntaje > mejor_puntaje:
            mejor, mejor_puntaje = p, puntaje
    return mejor if mejor_puntaje >= 1.2 else None


# ------------------------------------------------------------ base de datos


def db() -> sqlite3.Connection:
    con = sqlite3.connect(DB_PATH)
    con.execute(
        """CREATE TABLE IF NOT EXISTS movimientos (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            usuario INTEGER NOT NULL,
            tipo TEXT NOT NULL CHECK (tipo IN ('gasto', 'ingreso')),
            monto REAL NOT NULL,
            descripcion TEXT,
            fecha TEXT NOT NULL,
            categoria TEXT DEFAULT 'Otros'
        )"""
    )
    columnas = {c[1] for c in con.execute("PRAGMA table_info(movimientos)")}
    if "categoria" not in columnas:  # bases creadas con la versión anterior
        con.execute("ALTER TABLE movimientos ADD COLUMN categoria TEXT DEFAULT 'Otros'")
    return con


def guardar_movimiento(usuario: int, tipo: str, monto: float, categoria: str, descripcion: str) -> None:
    with db() as con:
        con.execute(
            "INSERT INTO movimientos (usuario, tipo, monto, categoria, descripcion, fecha) VALUES (?, ?, ?, ?, ?, ?)",
            (usuario, tipo, monto, categoria, descripcion, datetime.now().isoformat(timespec="seconds")),
        )


def movimientos_desde(usuario: int, desde: str) -> list[tuple]:
    with db() as con:
        return con.execute(
            "SELECT fecha, tipo, categoria, descripcion, monto FROM movimientos "
            "WHERE usuario = ? AND fecha >= ? ORDER BY fecha",
            (usuario, desde),
        ).fetchall()


def parsear_monto(texto: str) -> float | None:
    """Acepta 15000, 15.000, $15.000, 15,50, 1.234,56, 1,234.56."""
    t = texto.replace("$", "").replace(" ", "")
    if not re.fullmatch(r"[\d.,]+", t) or not re.search(r"\d", t):
        return None
    if "." in t and "," in t:
        decimal = "." if t.rfind(".") > t.rfind(",") else ","
        miles = "," if decimal == "." else "."
        t = t.replace(miles, "").replace(decimal, ".")
    elif "." in t or "," in t:
        sep = "." if "." in t else ","
        partes = t.split(sep)
        if len(partes) > 2 or len(partes[-1]) == 3:  # separador de miles
            t = t.replace(sep, "")
        else:
            t = t.replace(sep, ".")
    try:
        valor = float(t)
    except ValueError:
        return None
    return valor if valor > 0 else None


def formato(valor: float) -> str:
    entero = f"{valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return "$" + (entero[:-3] if entero.endswith(",00") else entero)


def icono(tipo: str, categoria: str) -> str:
    return CATEGORIAS_MOV[tipo].get(categoria, "📦")


# ------------------------------------------------------------------ teclados


def teclado_faq() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([[InlineKeyboardButton(n, callback_data=f"cat:{k}")]
                                 for k, n in CATEGORIAS_FAQ.items()])


def teclado_preguntas(categoria: str) -> InlineKeyboardMarkup:
    botones = [[InlineKeyboardButton(p["pregunta"], callback_data=f"q:{i}")]
               for i, p in enumerate(PREGUNTAS) if p["categoria"] == categoria]
    botones.append([InlineKeyboardButton("⬅️ Volver", callback_data="faq")])
    return InlineKeyboardMarkup(botones)


def teclado_volver(categoria: str) -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup([
        [InlineKeyboardButton("⬅️ Más preguntas", callback_data=f"cat:{categoria}")],
        [InlineKeyboardButton("📚 Todos los temas", callback_data="faq")],
    ])


def teclado_categorias_mov(tipo: str) -> InlineKeyboardMarkup:
    items = [InlineKeyboardButton(f"{e} {c}", callback_data=f"rcat:{c}")
             for c, e in CATEGORIAS_MOV[tipo].items()]
    filas = [items[i:i + 2] for i in range(0, len(items), 2)]
    filas.append([InlineKeyboardButton("✖️ Cancelar", callback_data="cancelar")])
    return InlineKeyboardMarkup(filas)


CANCELAR = InlineKeyboardMarkup([[InlineKeyboardButton("✖️ Cancelar", callback_data="cancelar")]])

# ------------------------------------------------------------ inicio y ayuda

AYUDA = (
    "Usa los botones de abajo 👇\n\n"
    "➖ / ➕ <b>Registrar</b> un gasto o ingreso: eliges categoría, escribes el monto y listo.\n"
    "📊 <b>Resumen</b> del mes actual.\n"
    "📥 📄 <b>Reportes</b> en Excel o Word con gráficas.\n"
    "❓ <b>Preguntas</b> sobre finanzas (también puedes escribirlas).\n"
    "🗑️ <b>Borrar</b> el último movimiento si te equivocaste."
)


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    nombre = html.escape(update.effective_user.first_name or "")
    saludo = f"¡Hola {nombre}! 👋 Soy <b>MIRA</b>, tu asistente de finanzas personales.\n\n{AYUDA}"
    if LOGO.exists():
        with open(LOGO, "rb") as foto:
            await update.message.reply_photo(foto, caption=saludo, parse_mode=ParseMode.HTML, reply_markup=MENU)
    else:
        await update.message.reply_text(saludo, parse_mode=ParseMode.HTML, reply_markup=MENU)


async def ayuda(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.message.reply_text(AYUDA, parse_mode=ParseMode.HTML, reply_markup=MENU)


async def preguntas(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    await update.message.reply_text("📚 Elige un tema:", reply_markup=teclado_faq())


# ------------------------------------------------- registro guiado con botones


async def inicio_gasto(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    return await _inicio_registro(update, context, "gasto")


async def inicio_ingreso(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    return await _inicio_registro(update, context, "ingreso")


async def _inicio_registro(update: Update, context: ContextTypes.DEFAULT_TYPE, tipo: str) -> int:
    context.user_data["reg"] = {"tipo": tipo}
    texto = "➖ <b>Nuevo gasto</b>" if tipo == "gasto" else "➕ <b>Nuevo ingreso</b>"
    await update.message.reply_text(f"{texto}\n¿En qué categoría?", parse_mode=ParseMode.HTML,
                                    reply_markup=teclado_categorias_mov(tipo))
    return CATEGORIA


async def elegir_categoria(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    query = update.callback_query
    await query.answer()
    reg = context.user_data.setdefault("reg", {"tipo": "gasto"})
    reg["categoria"] = query.data.split(":", 1)[1]
    await query.edit_message_text(
        f"{icono(reg['tipo'], reg['categoria'])} <b>{reg['categoria']}</b>\n\n"
        "💲 Escribe el monto (ej. <code>25000</code> o <code>25.000</code>):",
        parse_mode=ParseMode.HTML, reply_markup=CANCELAR,
    )
    return MONTO


async def categoria_por_texto(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    await update.message.reply_text("Elige la categoría con los botones de arriba 👆")
    return CATEGORIA


async def recibir_monto(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    monto = parsear_monto(update.message.text)
    if monto is None:
        await update.message.reply_text("No entendí el monto 🤔. Escribe solo el número, ej. 25000",
                                        reply_markup=CANCELAR)
        return MONTO
    context.user_data["reg"]["monto"] = monto
    await update.message.reply_text(
        f"Monto: <b>{formato(monto)}</b>\n\n✏️ Escribe una descripción corta (ej. <i>almuerzo</i>) "
        "o toca <b>Omitir</b>:",
        parse_mode=ParseMode.HTML,
        reply_markup=InlineKeyboardMarkup([[
            InlineKeyboardButton("⏭️ Omitir", callback_data="omitir"),
            InlineKeyboardButton("✖️ Cancelar", callback_data="cancelar"),
        ]]),
    )
    return DESCRIPCION


async def recibir_descripcion(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    return await _finalizar_registro(update, context, update.message.text.strip()[:100])


async def omitir_descripcion(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    await update.callback_query.answer()
    await update.callback_query.edit_message_reply_markup(None)
    return await _finalizar_registro(update, context, "")


async def _finalizar_registro(update: Update, context: ContextTypes.DEFAULT_TYPE, descripcion: str) -> int:
    reg = context.user_data.pop("reg", None)
    if not reg or "monto" not in reg:
        return ConversationHandler.END
    usuario = update.effective_user.id
    guardar_movimiento(usuario, reg["tipo"], reg["monto"], reg["categoria"], descripcion or reg["categoria"])
    filas = movimientos_desde(usuario, datetime.now().strftime("%Y-%m-01"))
    d = reportes.analizar(filas)
    simbolo = "🔴" if reg["tipo"] == "gasto" else "🟢"
    texto = (
        f"{simbolo} <b>{reg['tipo'].capitalize()} registrado</b>\n"
        f"{icono(reg['tipo'], reg['categoria'])} {reg['categoria']} · {formato(reg['monto'])}"
        + (f"\n📝 {html.escape(descripcion)}" if descripcion else "")
        + f"\n\nBalance del mes: <b>{'-' if d['balance'] < 0 else ''}{formato(abs(d['balance']))}</b>"
    )
    await update.effective_message.reply_text(texto, parse_mode=ParseMode.HTML, reply_markup=MENU)
    return ConversationHandler.END


async def cancelar(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data.pop("reg", None)
    if update.callback_query:
        await update.callback_query.answer()
        await update.callback_query.edit_message_text("✖️ Registro cancelado.")
    await update.effective_message.reply_text("¿Qué quieres hacer?", reply_markup=MENU)
    return ConversationHandler.END


async def reiniciar(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    context.user_data.pop("reg", None)
    await start(update, context)
    return ConversationHandler.END


async def boton_menu_durante_registro(update: Update, context: ContextTypes.DEFAULT_TYPE) -> int:
    """Si el usuario toca otro botón del menú a mitad de un registro, se cancela y se atiende el botón."""
    context.user_data.pop("reg", None)
    accion = ACCIONES_MENU.get(update.message.text)
    resultado = await accion(update, context) if accion else None
    return resultado if resultado is not None else ConversationHandler.END


# ----------------------------------------------------- resumen y movimientos


async def resumen(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    mes = datetime.now().strftime("%Y-%m")
    d = reportes.analizar(movimientos_desde(update.effective_user.id, f"{mes}-01"))
    texto = (
        f"📅 <b>Resumen de {reportes._mes(mes)}</b>\n\n"
        f"🟢 Ingresos: {formato(d['ingresos'])}\n"
        f"🔴 Gastos: {formato(d['gastos'])}\n"
        f"{'✅' if d['balance'] >= 0 else '⚠️'} Balance: <b>{'-' if d['balance'] < 0 else ''}"
        f"{formato(abs(d['balance']))}</b>"
    )
    if d["ingresos"] > 0:
        texto += f"\n\nHas gastado el {d['gastos'] / d['ingresos']:.0%} de tus ingresos."
    if d["categorias"]:
        texto += "\n\n<b>Gastos por categoría:</b>\n" + "\n".join(
            f"{icono('gasto', c)} {c}: {formato(v)} ({v / d['gastos']:.0%})" for c, v in d["categorias"]
        )
    texto += "\n\nDescarga el detalle con gráficas en 📥 Excel o 📄 Word."
    await update.message.reply_text(texto, parse_mode=ParseMode.HTML, reply_markup=MENU)


async def ultimos(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    with db() as con:
        filas = con.execute(
            "SELECT tipo, monto, categoria, descripcion, fecha FROM movimientos "
            "WHERE usuario = ? ORDER BY id DESC LIMIT 10",
            (update.effective_user.id,),
        ).fetchall()
    if not filas:
        await update.message.reply_text("Aún no tienes movimientos. Toca ➖ Registrar gasto para empezar.",
                                        reply_markup=MENU)
        return
    lineas = [
        f"{'🔴' if t == 'gasto' else '🟢'} {f[5:10]} {icono(t, c or 'Otros')} {formato(m)} – {html.escape(d or '')}"
        for t, m, c, d, f in filas
    ]
    await update.message.reply_text("<b>Tus últimos movimientos:</b>\n\n" + "\n".join(lineas),
                                    parse_mode=ParseMode.HTML, reply_markup=MENU)


async def borrar(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    with db() as con:
        fila = con.execute(
            "SELECT id, tipo, monto, categoria, descripcion FROM movimientos WHERE usuario = ? ORDER BY id DESC LIMIT 1",
            (update.effective_user.id,),
        ).fetchone()
    if not fila:
        await update.message.reply_text("No tienes movimientos para borrar.", reply_markup=MENU)
        return
    mid, tipo, monto, cat, desc = fila
    await update.message.reply_text(
        f"¿Borrar este {tipo}?\n{icono(tipo, cat or 'Otros')} {formato(monto)} – {html.escape(desc or '')}",
        reply_markup=InlineKeyboardMarkup([[
            InlineKeyboardButton("🗑️ Sí, borrar", callback_data=f"borrar:{mid}"),
            InlineKeyboardButton("No", callback_data="borrar:no"),
        ]]),
    )


async def confirmar_borrar(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    query = update.callback_query
    await query.answer()
    valor = query.data.split(":", 1)[1]
    if valor == "no":
        await query.edit_message_text("Listo, no se borró nada.")
        return
    with db() as con:
        borrados = con.execute("DELETE FROM movimientos WHERE id = ? AND usuario = ?",
                               (int(valor), update.effective_user.id)).rowcount
    await query.edit_message_text("🗑️ Movimiento eliminado." if borrados else "Ese movimiento ya no existe.")


# ------------------------------------------------------------------ reportes

PERIODOS = {"mes": "Este mes", "3m": "Últimos 3 meses", "anio": "Este año", "todo": "Todo"}


def fecha_desde(periodo: str) -> str:
    hoy = datetime.now()
    if periodo == "mes":
        return hoy.strftime("%Y-%m-01")
    if periodo == "3m":
        mes, anio = hoy.month - 2, hoy.year
        if mes <= 0:
            mes, anio = mes + 12, anio - 1
        return f"{anio}-{mes:02d}-01"
    if periodo == "anio":
        return f"{hoy.year}-01-01"
    return ""


async def pedir_reporte(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    formato_rep = "xlsx" if update.message.text == B_EXCEL else "docx"
    nombre = "Excel" if formato_rep == "xlsx" else "Word"
    botones = [[InlineKeyboardButton(t, callback_data=f"rep:{formato_rep}:{k}")] for k, t in PERIODOS.items()]
    await update.message.reply_text(f"¿De qué periodo quieres el reporte en {nombre}?",
                                    reply_markup=InlineKeyboardMarkup(botones))


async def enviar_reporte(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    query = update.callback_query
    await query.answer()
    _, formato_rep, periodo = query.data.split(":")
    filas = movimientos_desde(update.effective_user.id, fecha_desde(periodo))
    if not filas:
        await query.edit_message_text(f"No tienes movimientos en: {PERIODOS[periodo]}.")
        return
    await query.edit_message_text(f"⏳ Generando tu reporte ({PERIODOS[periodo]})...")
    await context.bot.send_chat_action(update.effective_chat.id, ChatAction.UPLOAD_DOCUMENT)
    generar = reportes.generar_excel if formato_rep == "xlsx" else reportes.generar_word
    archivo = await asyncio.to_thread(generar, filas, PERIODOS[periodo])
    nombre = f"MIRA_reporte_{periodo}_{datetime.now():%Y-%m-%d}.{formato_rep}"
    await query.message.reply_document(archivo, filename=nombre,
                                       caption=f"📊 Reporte {PERIODOS[periodo].lower()} · {len(filas)} movimientos",
                                       reply_markup=MENU)
    await query.edit_message_text(f"✅ Reporte listo ({PERIODOS[periodo]}).")


# --------------------------------------------- comandos rápidos (opcionales)


async def comando_registro(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    """/gasto 25000 almuerzo  – atajo para quien prefiera escribir."""
    tipo = "gasto" if update.message.text.startswith("/gasto") else "ingreso"
    monto = parsear_monto(context.args[0]) if context.args else None
    if monto is None:
        await (inicio_gasto if tipo == "gasto" else inicio_ingreso)(update, context)
        return
    descripcion = " ".join(context.args[1:]) or "Otros"
    guardar_movimiento(update.effective_user.id, tipo, monto, "Otros", descripcion)
    await update.message.reply_text(f"{'🔴' if tipo == 'gasto' else '🟢'} {tipo.capitalize()} registrado: "
                                    f"{formato(monto)} – {descripcion}", reply_markup=MENU)


# ------------------------------------------------------- FAQ y texto libre


async def botones_faq(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    query = update.callback_query
    await query.answer()
    data = query.data
    if data == "faq":
        await query.edit_message_text("📚 Elige un tema:", reply_markup=teclado_faq())
    elif data.startswith("cat:"):
        clave = data[4:]
        await query.edit_message_text(f"{CATEGORIAS_FAQ.get(clave, 'Preguntas')}\nElige una pregunta:",
                                      reply_markup=teclado_preguntas(clave))
    elif data.startswith("q:"):
        p = PREGUNTAS[int(data[2:])]
        await query.edit_message_text(f"❓ {p['pregunta']}\n\n{p['respuesta']}",
                                      reply_markup=teclado_volver(p["categoria"]))


async def texto_libre(update: Update, context: ContextTypes.DEFAULT_TYPE) -> None:
    p = buscar_pregunta(update.message.text)
    if p:
        await update.message.reply_text(f"❓ {p['pregunta']}\n\n{p['respuesta']}",
                                        reply_markup=teclado_volver(p["categoria"]))
    else:
        await update.message.reply_text(
            "No encontré una respuesta para eso 🤔.\nUsa los botones del menú o mira los temas:",
            reply_markup=teclado_faq(),
        )


async def error(update: object, context: ContextTypes.DEFAULT_TYPE) -> None:
    log.error("Error procesando una actualización", exc_info=context.error)


ACCIONES_MENU = {
    B_GASTO: inicio_gasto, B_INGRESO: inicio_ingreso, B_RESUMEN: resumen, B_ULTIMOS: ultimos,
    B_EXCEL: pedir_reporte, B_WORD: pedir_reporte, B_FAQ: preguntas, B_BORRAR: borrar,
}


def main() -> None:
    if not TOKEN:
        raise SystemExit("Falta TELEGRAM_TOKEN en el archivo .env")
    db().close()
    espera = 5
    while True:
        try:
            crear_app().run_polling(allowed_updates=Update.ALL_TYPES, close_loop=False)
            return  # se detuvo normalmente con Ctrl+C
        except (NetworkError, TimedOut) as e:
            log.warning("No se pudo conectar con Telegram (%s). Reintentando en %s s... "
                        "Si sigue fallando, revisa tu internet o activa la VPN.", type(e).__name__, espera)
            time.sleep(espera)
            espera = min(espera * 2, 60)


def crear_app() -> Application:
    builder = (Application.builder().token(TOKEN)
               .connect_timeout(20).read_timeout(30).write_timeout(30)
               .get_updates_connect_timeout(20).get_updates_read_timeout(40))
    proxy = os.getenv("TELEGRAM_PROXY")  # opcional, ej. http://127.0.0.1:8080
    if proxy:
        builder = builder.proxy(proxy).get_updates_proxy(proxy)
    app = builder.build()

    es_boton = filters.Text(BOTONES_MENU)
    texto_normal = filters.TEXT & ~filters.COMMAND & ~es_boton
    registro = ConversationHandler(
        entry_points=[
            MessageHandler(filters.Text([B_GASTO]), inicio_gasto),
            MessageHandler(filters.Text([B_INGRESO]), inicio_ingreso),
        ],
        states={
            CATEGORIA: [CallbackQueryHandler(elegir_categoria, pattern=r"^rcat:"),
                        MessageHandler(texto_normal, categoria_por_texto)],
            MONTO: [MessageHandler(texto_normal, recibir_monto)],
            DESCRIPCION: [MessageHandler(texto_normal, recibir_descripcion),
                          CallbackQueryHandler(omitir_descripcion, pattern=r"^omitir$")],
        },
        fallbacks=[
            CallbackQueryHandler(cancelar, pattern=r"^cancelar$"),
            CommandHandler("cancelar", cancelar),
            MessageHandler(es_boton, boton_menu_durante_registro),
            CommandHandler("start", reiniciar),
        ],
    )
    app.add_handler(registro)
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler(["ayuda", "help"], ayuda))
    app.add_handler(CommandHandler("menu", ayuda))
    app.add_handler(CommandHandler(["gasto", "ingreso"], comando_registro))
    app.add_handler(CommandHandler("resumen", resumen))
    app.add_handler(MessageHandler(filters.Text([B_RESUMEN]), resumen))
    app.add_handler(MessageHandler(filters.Text([B_ULTIMOS]), ultimos))
    app.add_handler(MessageHandler(filters.Text([B_EXCEL, B_WORD]), pedir_reporte))
    app.add_handler(MessageHandler(filters.Text([B_FAQ]), preguntas))
    app.add_handler(MessageHandler(filters.Text([B_BORRAR]), borrar))
    app.add_handler(CallbackQueryHandler(enviar_reporte, pattern=r"^rep:"))
    app.add_handler(CallbackQueryHandler(confirmar_borrar, pattern=r"^borrar:"))
    app.add_handler(CallbackQueryHandler(botones_faq, pattern=r"^(faq|cat:|q:)"))
    app.add_handler(CallbackQueryHandler(cancelar, pattern=r"^cancelar$"))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, texto_libre))
    app.add_error_handler(error)
    log.info("Conectando con Telegram... (Ctrl+C para detener)")
    app.post_init = _conectado
    return app


async def _conectado(app: Application) -> None:
    log.info("✅ Bot MIRA conectado como @%s. Ya puedes usarlo en Telegram.", app.bot.username)


if __name__ == "__main__":
    main()
