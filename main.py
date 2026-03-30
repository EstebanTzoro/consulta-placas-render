from fastapi import FastAPI, Form, HTTPException
from fastapi.responses import HTMLResponse, StreamingResponse
from io import BytesIO
from typing import List
import pandas as pd
import requests
import os
from datetime import datetime, timezone

app = FastAPI(title="Consulta de placas", version="1.0.0")

# =========================
# CONFIG
# =========================
PUBLIC_FILE_URL = os.getenv(
    "PUBLIC_FILE_URL",
    "https://celerix-my.sharepoint.com/:x:/p/jtcaraballo/IQAXf4-pTrwPQL8iYQYuwvG7Aaj1bjCB4IbSi03p97qlftI?download=1"
)

SHEET_ESCENARIOS = "Escenarios"
SHEET_ESCENARIOS_ESTIMADO = "Escenarios_Estimado"
COL_PLACA = "Placa"

CACHE_DATA = None
CACHE_TS = None
CACHE_MINUTES = 5


# =========================
# HELPERS
# =========================
def normalizar_placa(valor) -> str:
    if pd.isna(valor):
        return ""
    return (
        str(valor)
        .upper()
        .strip()
        .replace(" ", "")
        .replace("-", "")
    )


def descargar_excel_publico() -> bytes:
    try:
        r = requests.get(PUBLIC_FILE_URL, timeout=120, allow_redirects=True)
    except Exception as e:
        raise HTTPException(
            status_code=500,
            detail=f"No se pudo conectar al archivo público: {e}"
        )

    if r.status_code != 200:
        raise HTTPException(
            status_code=500,
            detail=f"No se pudo descargar el archivo. Status={r.status_code}"
        )

    content_type = (r.headers.get("content-type") or "").lower()

    if "text/html" in content_type:
        raise HTTPException(
            status_code=500,
            detail="El link devolvió HTML en vez del Excel."
        )

    return r.content


def cargar_datos(force_refresh: bool = False) -> dict:
    global CACHE_DATA, CACHE_TS

    now = datetime.now(timezone.utc)

    if (
        not force_refresh
        and CACHE_DATA is not None
        and CACHE_TS is not None
        and (now - CACHE_TS).total_seconds() < CACHE_MINUTES * 60
    ):
        return {
            "escenarios": CACHE_DATA["escenarios"].copy(),
            "escenarios_estimado": CACHE_DATA["escenarios_estimado"].copy(),
        }

    content = descargar_excel_publico()

    try:
        xls = pd.ExcelFile(BytesIO(content))
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"No se pudo abrir el Excel: {e}")

    hojas_requeridas = [SHEET_ESCENARIOS, SHEET_ESCENARIOS_ESTIMADO]
    faltantes = [h for h in hojas_requeridas if h not in xls.sheet_names]
    if faltantes:
        raise HTTPException(
            status_code=500,
            detail=f"Faltan hojas requeridas en el Excel: {faltantes}. Hojas encontradas: {xls.sheet_names}"
        )

    try:
        df_escenarios = pd.read_excel(xls, sheet_name=SHEET_ESCENARIOS)
        df_estimado = pd.read_excel(xls, sheet_name=SHEET_ESCENARIOS_ESTIMADO)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"No se pudieron leer las hojas: {e}")

    for nombre_hoja, df in [
        (SHEET_ESCENARIOS, df_escenarios),
        (SHEET_ESCENARIOS_ESTIMADO, df_estimado),
    ]:
        if COL_PLACA not in df.columns:
            raise HTTPException(
                status_code=500,
                detail=f"La columna '{COL_PLACA}' no existe en la hoja '{nombre_hoja}'. Columnas encontradas: {list(df.columns)}"
            )

    df_escenarios[COL_PLACA] = df_escenarios[COL_PLACA].apply(normalizar_placa)
    df_estimado[COL_PLACA] = df_estimado[COL_PLACA].apply(normalizar_placa)

    if "Rank" not in df_estimado.columns:
        raise HTTPException(
            status_code=500,
            detail=f"La columna 'Rank' no existe en la hoja '{SHEET_ESCENARIOS_ESTIMADO}'. Columnas encontradas: {list(df_estimado.columns)}"
        )

    df_estimado["Rank"] = pd.to_numeric(df_estimado["Rank"], errors="coerce")
    df_estimado = df_estimado[df_estimado["Rank"].isin([1, 2, 3, 4, 5])].copy()


    CACHE_DATA = {
        "escenarios": df_escenarios.copy(),
        "escenarios_estimado": df_estimado.copy(),
    }
    CACHE_TS = now

    return {
        "escenarios": df_escenarios,
        "escenarios_estimado": df_estimado,
    }


def limpiar_placas(texto: str) -> List[str]:
    raw = texto.replace(";", "\n").replace(",", "\n").splitlines()
    placas = [normalizar_placa(x) for x in raw if normalizar_placa(x)]

    seen = set()
    salida = []
    for p in placas:
        if p not in seen:
            seen.add(p)
            salida.append(p)
    return salida


def filtrar_por_placas(df: pd.DataFrame, placas: List[str]) -> pd.DataFrame:
    return df[df[COL_PLACA].isin(placas)].copy()


def excel_resultado_bytes(df_escenarios: pd.DataFrame, df_estimado: pd.DataFrame) -> BytesIO:
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_escenarios.to_excel(writer, sheet_name=SHEET_ESCENARIOS, index=False)
        df_estimado.to_excel(writer, sheet_name=SHEET_ESCENARIOS_ESTIMADO, index=False)

    output.seek(0)
    return output


def html_escape(texto: str) -> str:
    return (
        str(texto)
        .replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&#39;")
    )


# =========================
# SHARED CSS + FONTS
# =========================
BASE_STYLES = """
@import url('https://fonts.googleapis.com/css2?family=Cormorant+Garamond:ital,wght@0,300;0,400;0,500;0,600;1,300;1,400&family=Instrument+Sans:wght@300;400;500;600&family=Geist+Mono:wght@300;400;500&display=swap');

*, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

:root {
    --bg: #0a0a0f;
    --surface: #12121a;
    --surface2: #1a1a26;
    --border: rgba(255,255,255,0.07);
    --border-hover: rgba(255,255,255,0.15);
    --accent: #c8f135;
    --accent-dim: rgba(200,241,53,0.12);
    --text: #f0f0f0;
    --text-muted: #666680;
    --text-soft: #9898b0;
    --error-bg: rgba(255,80,80,0.08);
    --error-border: rgba(255,80,80,0.25);
    --error-text: #ff7070;
    --success-bg: rgba(200,241,53,0.07);
    --radius: 16px;
    --radius-sm: 10px;
}

html { scroll-behavior: smooth; }

body {
    font-family: 'Instrument Sans', sans-serif;
    background: var(--bg);
    color: var(--text);
    min-height: 100vh;
    line-height: 1.6;
    font-size: 15px;
    -webkit-font-smoothing: antialiased;
}

/* Noise texture overlay */
body::before {
    content: '';
    position: fixed;
    inset: 0;
    background-image: url("data:image/svg+xml,%3Csvg viewBox='0 0 256 256' xmlns='http://www.w3.org/2000/svg'%3E%3Cfilter id='n'%3E%3CfeTurbulence type='fractalNoise' baseFrequency='0.9' numOctaves='4' stitchTiles='stitch'/%3E%3C/filter%3E%3Crect width='100%25' height='100%25' filter='url(%23n)' opacity='0.03'/%3E%3C/svg%3E");
    pointer-events: none;
    z-index: 0;
    opacity: 0.4;
}

.page {
    position: relative;
    z-index: 1;
    max-width: 860px;
    margin: 0 auto;
    padding: 64px 28px 100px;
}

/* Header */
.header {
    margin-bottom: 52px;
    animation: fadeUp 0.5s ease both;
}

.header-eyebrow {
    font-family: 'Geist Mono', monospace;
    font-size: 10px;
    letter-spacing: 0.2em;
    text-transform: uppercase;
    color: var(--accent);
    margin-bottom: 18px;
    display: flex;
    align-items: center;
    gap: 10px;
    font-weight: 300;
}

.header-eyebrow::before {
    content: '';
    display: block;
    width: 28px;
    height: 1px;
    background: var(--accent);
}

.header h1 {
    font-family: 'Cormorant Garamond', serif;
    font-size: clamp(44px, 6vw, 72px);
    font-weight: 300;
    letter-spacing: -0.01em;
    line-height: 1.05;
    color: var(--text);
}

.header h1 em {
    font-style: italic;
    font-weight: 300;
    color: var(--accent);
}

.header h1 span {
    color: var(--accent);
    font-style: italic;
}

.header-desc {
    margin-top: 18px;
    color: var(--text-soft);
    font-size: 14px;
    font-weight: 300;
    max-width: 480px;
    line-height: 1.8;
    letter-spacing: 0.01em;
}

/* Cards */
.card {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    padding: 32px;
    animation: fadeUp 0.5s ease both;
    transition: border-color 0.2s;
}

.card:hover {
    border-color: var(--border-hover);
}

.card-label {
    font-family: 'Geist Mono', monospace;
    font-size: 10px;
    letter-spacing: 0.18em;
    text-transform: uppercase;
    color: var(--text-muted);
    margin-bottom: 20px;
    display: flex;
    align-items: center;
    gap: 8px;
    font-weight: 300;
}

.card-label::after {
    content: '';
    flex: 1;
    height: 1px;
    background: var(--border);
}

/* Textarea */
textarea {
    width: 100%;
    height: 200px;
    background: var(--surface2);
    border: 1px solid var(--border);
    border-radius: var(--radius-sm);
    color: var(--text);
    font-family: 'Geist Mono', monospace;
    font-size: 13px;
    line-height: 1.9;
    padding: 18px 20px;
    resize: vertical;
    outline: none;
    transition: border-color 0.2s, box-shadow 0.2s;
    letter-spacing: 0.05em;
    font-weight: 300;
}

textarea::placeholder { color: var(--text-muted); }

textarea:focus {
    border-color: var(--accent);
    box-shadow: 0 0 0 3px var(--accent-dim);
}

/* Button */
.btn {
    display: inline-flex;
    align-items: center;
    gap: 10px;
    background: var(--accent);
    color: #0a0a0f;
    border: none;
    padding: 13px 24px;
    font-family: 'Instrument Sans', sans-serif;
    font-size: 13px;
    font-weight: 500;
    letter-spacing: 0.06em;
    text-transform: uppercase;
    border-radius: var(--radius-sm);
    cursor: pointer;
    text-decoration: none;
    transition: opacity 0.15s, transform 0.15s, box-shadow 0.15s;
    box-shadow: 0 4px 24px rgba(200,241,53,0.2);
}

.btn:hover {
    opacity: 0.9;
    transform: translateY(-1px);
    box-shadow: 0 8px 32px rgba(200,241,53,0.3);
}

.btn:active { transform: translateY(0); }

.btn-icon {
    width: 16px;
    height: 16px;
    flex-shrink: 0;
}

.btn-ghost {
    background: transparent;
    color: var(--text-soft);
    border: 1px solid var(--border);
    box-shadow: none;
    font-weight: 500;
}

.btn-ghost:hover {
    border-color: var(--border-hover);
    color: var(--text);
    box-shadow: none;
    background: var(--surface2);
}

/* Hint */
.hint {
    margin-top: 14px;
    font-size: 12px;
    color: var(--text-muted);
    font-family: 'Geist Mono', monospace;
    font-weight: 300;
    letter-spacing: 0.04em;
}

.hint a {
    color: var(--accent);
    text-decoration: none;
}

.hint a:hover { text-decoration: underline; }

/* Stats grid */
.stats-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(160px, 1fr));
    gap: 14px;
    margin-bottom: 28px;
}

.stat-item {
    background: var(--surface2);
    border: 1px solid var(--border);
    border-radius: var(--radius-sm);
    padding: 20px 18px;
    transition: border-color 0.2s;
}

.stat-item:hover { border-color: var(--border-hover); }

.stat-value {
    font-family: 'Cormorant Garamond', serif;
    font-size: 42px;
    font-weight: 300;
    color: var(--accent);
    line-height: 1;
    margin-bottom: 8px;
    letter-spacing: -0.02em;
}

.stat-label {
    font-family: 'Geist Mono', monospace;
    font-size: 9px;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    color: var(--text-muted);
    line-height: 1.5;
    font-weight: 300;
}

/* Cols */
.cols {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 16px;
    margin-top: 20px;
}

@media (max-width: 640px) {
    .cols { grid-template-columns: 1fr; }
    .stats-grid { grid-template-columns: 1fr 1fr; }
}

/* Placa tags */
.placa-list {
    display: flex;
    flex-wrap: wrap;
    gap: 7px;
    margin-top: 8px;
}

.placa-tag {
    font-family: 'Geist Mono', monospace;
    font-size: 11px;
    letter-spacing: 0.1em;
    background: var(--surface2);
    border: 1px solid var(--border);
    padding: 5px 12px;
    border-radius: 6px;
    color: var(--text-soft);
    transition: border-color 0.15s, color 0.15s;
    font-weight: 300;
}

.placa-tag:hover {
    border-color: var(--accent);
    color: var(--accent);
}

.placa-empty {
    font-family: 'Geist Mono', monospace;
    font-size: 12px;
    color: var(--text-muted);
    padding: 8px 0;
    font-weight: 300;
    font-style: italic;
}

/* Error / success banners */
.banner {
    display: flex;
    align-items: flex-start;
    gap: 14px;
    padding: 18px 22px;
    border-radius: var(--radius-sm);
    margin-bottom: 24px;
    font-size: 14px;
}

.banner-error {
    background: var(--error-bg);
    border: 1px solid var(--error-border);
    color: var(--error-text);
}

.banner-success {
    background: var(--success-bg);
    border: 1px solid rgba(200,241,53,0.2);
    color: var(--accent);
}

.banner-icon {
    flex-shrink: 0;
    margin-top: 1px;
}

/* Divider */
.divider {
    height: 1px;
    background: var(--border);
    margin: 28px 0;
}

/* Back link */
.back-link {
    display: inline-flex;
    align-items: center;
    gap: 8px;
    font-family: 'Geist Mono', monospace;
    font-size: 11px;
    letter-spacing: 0.1em;
    color: var(--text-muted);
    text-decoration: none;
    margin-bottom: 40px;
    transition: color 0.2s;
    font-weight: 300;
}

.back-link:hover { color: var(--accent); }

/* Animations */
@keyframes fadeUp {
    from { opacity: 0; transform: translateY(16px); }
    to   { opacity: 1; transform: translateY(0); }
}

.card:nth-child(2) { animation-delay: 0.08s; }
.card:nth-child(3) { animation-delay: 0.16s; }

/* Actions row */
.actions {
    display: flex;
    align-items: center;
    gap: 12px;
    margin-top: 22px;
    flex-wrap: wrap;
}

/* Card title */
.card-title {
    font-family: 'Cormorant Garamond', serif;
    font-size: 22px;
    font-weight: 300;
    margin-bottom: 16px;
    color: var(--text);
    letter-spacing: 0.01em;
}

/* Loading state */
.btn-loading {
    pointer-events: none;
    opacity: 0.7;
}
"""

LOADING_JS = """
<script>
document.addEventListener('DOMContentLoaded', function() {
    const forms = document.querySelectorAll('form');
    forms.forEach(function(form) {
        form.addEventListener('submit', function() {
            const btn = form.querySelector('button[type="submit"]');
            if (btn) {
                btn.classList.add('btn-loading');
                btn.innerHTML = '<svg class="btn-icon spin" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M12 2v4M12 18v4M4.93 4.93l2.83 2.83M16.24 16.24l2.83 2.83M2 12h4M18 12h4M4.93 19.07l2.83-2.83M16.24 7.76l2.83-2.83"/></svg> Consultando...';
            }
        });
    });
});
</script>
<style>
@keyframes spin { to { transform: rotate(360deg); } }
.spin { animation: spin 0.8s linear infinite; }
</style>
"""


# =========================
# RUTAS
# =========================
@app.get("/", response_class=HTMLResponse)
def home():
    return f"""<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Consulta de Placas</title>
    <style>{BASE_STYLES}</style>
</head>
<body>
<div class="page">

    <header class="header">
        <div class="header-eyebrow">Sistema de consulta</div>
        <h1>Consulta de <span>Placas</span></h1>
        <p class="header-desc">
            Ingresa una o varias placas para obtener los registros filtrados
            de las hojas <strong>Escenarios</strong> y <strong>Escenarios_Estimado</strong> en formato Excel.
        </p>
    </header>

    <div class="card">
        <div class="card-label">Placas a consultar</div>
        <form action="/consultar-form" method="post">
            <textarea
                name="placas_texto"
                placeholder="ABC123&#10;DEF456&#10;GHI789&#10;&#10;Separa por coma, punto y coma o salto de línea"
                autofocus
            ></textarea>
            <div class="actions">
                <button class="btn" type="submit">
                    <svg class="btn-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><circle cx="11" cy="11" r="8"/><path d="m21 21-4.35-4.35"/></svg>
                    Consultar placas
                </button>
                <span class="hint">También disponible en <a href="/docs">/docs</a></span>
            </div>
        </form>
    </div>

</div>
{LOADING_JS}
</body>
</html>"""


@app.get("/health")
def health():
    return {"ok": True}


@app.post("/refresh-cache")
def refresh_cache():
    data = cargar_datos(force_refresh=True)
    return {
        "ok": True,
        "rows_escenarios": len(data["escenarios"]),
        "rows_escenarios_estimado": len(data["escenarios_estimado"]),
        "placas_unicas_escenarios": int(data["escenarios"][COL_PLACA].nunique()),
        "placas_unicas_escenarios_estimado": int(data["escenarios_estimado"][COL_PLACA].nunique()),
    }


@app.post("/descargar")
def descargar_excel(placas_texto: str = Form(...)):
    placas_lista = limpiar_placas(placas_texto)

    if not placas_lista:
        raise HTTPException(status_code=400, detail="Debes enviar al menos una placa.")

    data = cargar_datos()

    df_escenarios_filtrado = filtrar_por_placas(data["escenarios"], placas_lista)
    df_estimado_filtrado = filtrar_por_placas(data["escenarios_estimado"], placas_lista)

    if df_escenarios_filtrado.empty and df_estimado_filtrado.empty:
        raise HTTPException(
            status_code=404,
            detail="No se encontraron registros para las placas consultadas en ninguna de las dos hojas."
        )

    excel_io = excel_resultado_bytes(df_escenarios_filtrado, df_estimado_filtrado)

    nombre = f"consulta_placas_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    headers = {"Content-Disposition": f'attachment; filename="{nombre}"'}

    return StreamingResponse(
        excel_io,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers
    )


@app.post("/consultar-form", response_class=HTMLResponse)
def consultar_form(placas_texto: str = Form(...)):
    placas = limpiar_placas(placas_texto)

    if not placas:
        return f"""<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Consulta de Placas</title>
    <style>{BASE_STYLES}</style>
</head>
<body>
<div class="page">
    <a class="back-link" href="/">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M19 12H5M12 5l-7 7 7 7"/></svg>
        Volver al inicio
    </a>
    <div class="banner banner-error">
        <svg class="banner-icon" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>
        Debes ingresar al menos una placa para realizar la consulta.
    </div>
    <div class="card">
        <div class="card-label">Reintentar consulta</div>
        <form action="/consultar-form" method="post">
            <textarea name="placas_texto" placeholder="ABC123&#10;DEF456&#10;GHI789" autofocus></textarea>
            <div class="actions">
                <button class="btn" type="submit">
                    <svg class="btn-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><circle cx="11" cy="11" r="8"/><path d="m21 21-4.35-4.35"/></svg>
                    Consultar placas
                </button>
            </div>
        </form>
    </div>
</div>
{LOADING_JS}
</body>
</html>"""

    data = cargar_datos()

    df_escenarios_filtrado = filtrar_por_placas(data["escenarios"], placas)
    df_estimado_filtrado = filtrar_por_placas(data["escenarios_estimado"], placas)

    placas_encontradas_esc = sorted(df_escenarios_filtrado[COL_PLACA].dropna().unique().tolist())
    placas_encontradas_est = sorted(df_estimado_filtrado[COL_PLACA].dropna().unique().tolist())

    total_placas_consultadas = len(placas)
    total_registros_esc = len(df_escenarios_filtrado)
    total_registros_est = len(df_estimado_filtrado)
    total_placas_esc = len(placas_encontradas_esc)
    total_placas_est = len(placas_encontradas_est)

    placas_query = "\n".join(placas)

    # Build placa tag lists
    if placas_encontradas_esc:
        placas_esc_html = '<div class="placa-list">' + "".join(
            f'<span class="placa-tag">{html_escape(p)}</span>' for p in placas_encontradas_esc
        ) + '</div>'
    else:
        placas_esc_html = '<p class="placa-empty">— Sin resultados en esta hoja</p>'

    if placas_encontradas_est:
        placas_est_html = '<div class="placa-list">' + "".join(
            f'<span class="placa-tag">{html_escape(p)}</span>' for p in placas_encontradas_est
        ) + '</div>'
    else:
        placas_est_html = '<p class="placa-empty">— Sin resultados en esta hoja</p>'

    no_results = total_registros_esc == 0 and total_registros_est == 0

    result_banner = ""
    download_section = ""

    if no_results:
        result_banner = """
        <div class="banner banner-error">
            <svg class="banner-icon" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><circle cx="12" cy="12" r="10"/><line x1="12" y1="8" x2="12" y2="12"/><line x1="12" y1="16" x2="12.01" y2="16"/></svg>
            No se encontraron registros para las placas consultadas en ninguna de las dos hojas.
        </div>"""
    else:
        result_banner = """
        <div class="banner banner-success">
            <svg class="banner-icon" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"><polyline points="20 6 9 17 4 12"/></svg>
            Consulta completada. El archivo Excel está listo para descargar.
        </div>"""
        download_section = f"""
        <form action="/descargar" method="post" style="margin:0;">
            <textarea name="placas_texto" style="display:none;">{html_escape(placas_query)}</textarea>
            <button class="btn" type="submit">
                <svg class="btn-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>
                Descargar Excel filtrado
            </button>
        </form>"""

    return f"""<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Resultado — Consulta de Placas</title>
    <style>{BASE_STYLES}</style>
</head>
<body>
<div class="page">

    <a class="back-link" href="/">
        <svg width="14" height="14" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><path d="M19 12H5M12 5l-7 7 7 7"/></svg>
        Nueva consulta
    </a>

    <header class="header" style="margin-bottom: 32px;">
        <div class="header-eyebrow">Resultado</div>
        <h1>Resumen de <span>búsqueda</span></h1>
    </header>

    {result_banner}

    <!-- Stats -->
    <div class="stats-grid">
        <div class="stat-item">
            <div class="stat-value">{total_placas_consultadas}</div>
            <div class="stat-label">Placas consultadas</div>
        </div>
        <div class="stat-item">
            <div class="stat-value">{total_placas_esc}</div>
            <div class="stat-label">Encontradas en Escenarios</div>
        </div>
        <div class="stat-item">
            <div class="stat-value">{total_placas_est}</div>
            <div class="stat-label">Encontradas en Est. Estimado</div>
        </div>
        <div class="stat-item">
            <div class="stat-value">{total_registros_esc + total_registros_est}</div>
            <div class="stat-label">Total registros encontrados</div>
        </div>
    </div>

    <!-- Download CTA -->
    {f'<div class="actions" style="margin-bottom: 28px;">{download_section}</div>' if download_section else ''}

    <!-- Placa detail cards -->
    <div class="cols">
        <div class="card">
            <div class="card-label">Escenarios</div>
            <div class="card-title">{total_placas_esc} placa{"s" if total_placas_esc != 1 else ""} hallada{"s" if total_placas_esc != 1 else ""}</div>
            {placas_esc_html}
        </div>
        <div class="card">
            <div class="card-label">Escenarios Estimado</div>
            <div class="card-title">{total_placas_est} placa{"s" if total_placas_est != 1 else ""} hallada{"s" if total_placas_est != 1 else ""}</div>
            {placas_est_html}
        </div>
    </div>

    <!-- New query -->
    <div class="card" style="margin-top: 20px;">
        <div class="card-label">Nueva consulta</div>
        <form action="/consultar-form" method="post">
            <textarea name="placas_texto">{html_escape(placas_query)}</textarea>
            <div class="actions">
                <button class="btn" type="submit">
                    <svg class="btn-icon" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5"><circle cx="11" cy="11" r="8"/><path d="m21 21-4.35-4.35"/></svg>
                    Consultar de nuevo
                </button>
            </div>
        </form>
    </div>

</div>
{LOADING_JS}
</body>
</html>"""