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
# RUTAS
# =========================
@app.get("/", response_class=HTMLResponse)
def home():
    return """
    <html>
        <head>
            <title>Consulta de placas</title>
            <style>
                body { font-family: Arial, sans-serif; max-width: 950px; margin: 40px auto; padding: 0 20px; background: #ffffff; color: #111827; }
                h1 { margin-bottom: 8px; }
                p { color: #444; line-height: 1.5; }
                textarea { width: 100%; height: 220px; font-size: 14px; padding: 12px; border: 1px solid #d1d5db; border-radius: 10px; }
                button { background: #111827; color: white; border: none; padding: 12px 18px; font-size: 14px; border-radius: 8px; cursor: pointer; }
                button:hover { opacity: 0.92; }
                .box { background: #f9fafb; border: 1px solid #e5e7eb; padding: 20px; border-radius: 14px; }
                .hint { font-size: 13px; color: #6b7280; margin-top: 12px; }
                .summary { margin-top: 22px; padding: 18px; border-radius: 14px; background: #f3f4f6; border: 1px solid #e5e7eb; }
                .kpi { margin: 8px 0; font-size: 15px; }
                .download-btn { display: inline-block; margin-top: 18px; background: #111827; color: white; text-decoration: none; padding: 12px 18px; border-radius: 8px; }
                .download-btn:hover { opacity: 0.92; }
                .muted { color: #6b7280; }
                .error-box { background: #fef2f2; border: 1px solid #fecaca; color: #991b1b; padding: 16px; border-radius: 12px; }
            </style>
        </head>
        <body>
            <h1>Consulta de placas</h1>
            <p>
                Pega una o varias placas separadas por coma, punto y coma o salto de línea.
                El sistema consultará el archivo y te devolverá un Excel con las hojas
                <b>Escenarios</b> y <b>Escenarios_Estimado</b>, filtradas por las placas consultadas.
            </p>

            <div class="box">
                <form action="/consultar-form" method="post">
                    <textarea name="placas_texto" placeholder="ABC123
DEF456
GHI789"></textarea>
                    <br><br>
                    <button type="submit">Consultar placas</button>
                </form>
                <div class="hint">
                    También puedes probar la API en <a href="/docs">/docs</a>
                </div>
            </div>
        </body>
    </html>
    """


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


@app.get("/descargar")
def descargar_excel(placas: str):
    placas_lista = limpiar_placas(placas)

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
        return """
        <html>
            <head><title>Consulta de placas</title></head>
            <body style="font-family: Arial, sans-serif; max-width: 900px; margin: 40px auto; padding: 0 20px;">
                <div class="error-box" style="background:#fef2f2;border:1px solid #fecaca;color:#991b1b;padding:16px;border-radius:12px;">
                    Debes ingresar al menos una placa.
                </div>
                <p><a href="/">Volver</a></p>
            </body>
        </html>
        """

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

    placas_esc_html = "<br>".join(html_escape(p) for p in placas_encontradas_esc) if placas_encontradas_esc else "<span class='muted'>No se encontraron placas en esta hoja.</span>"
    placas_est_html = "<br>".join(html_escape(p) for p in placas_encontradas_est) if placas_encontradas_est else "<span class='muted'>No se encontraron placas en esta hoja.</span>"

    if total_registros_esc == 0 and total_registros_est == 0:
        return f"""
        <html>
            <head>
                <title>Resultado de consulta</title>
                <style>
                    body {{ font-family: Arial, sans-serif; max-width: 950px; margin: 40px auto; padding: 0 20px; }}
                    .box {{ background: #f9fafb; border: 1px solid #e5e7eb; padding: 20px; border-radius: 14px; }}
                    .error-box {{ background:#fef2f2; border:1px solid #fecaca; color:#991b1b; padding:16px; border-radius:12px; }}
                    .kpi {{ margin: 8px 0; }}
                    .muted {{ color:#6b7280; }}
                    textarea {{ width: 100%; height: 180px; font-size: 14px; padding: 12px; border: 1px solid #d1d5db; border-radius: 10px; }}
                    button {{ background: #111827; color: white; border: none; padding: 12px 18px; font-size: 14px; border-radius: 8px; cursor: pointer; }}
                </style>
            </head>
            <body>
                <h1>Resultado de consulta</h1>

                <div class="error-box">
                    No se encontraron registros para las placas consultadas en ninguna de las dos hojas.
                </div>

                <div class="box" style="margin-top:20px;">
                    <div class="kpi"><b>Placas consultadas:</b> {total_placas_consultadas}</div>
                    <div class="kpi"><b>Placas encontradas en Escenarios:</b> {total_placas_esc}</div>
                    <div class="kpi"><b>Placas encontradas en Escenarios_Estimado:</b> {total_placas_est}</div>
                    <div class="kpi"><b>Registros encontrados en Escenarios:</b> {total_registros_esc}</div>
                    <div class="kpi"><b>Registros encontrados en Escenarios_Estimado:</b> {total_registros_est}</div>
                </div>

                <div class="box" style="margin-top:20px;">
                    <h3>Consultar de nuevo</h3>
                    <form action="/consultar-form" method="post">
                        <textarea name="placas_texto">{html_escape(placas_query)}</textarea>
                        <br><br>
                        <button type="submit">Consultar placas</button>
                    </form>
                </div>
            </body>
        </html>
        """

    return f"""
    <html>
        <head>
            <title>Resultado de consulta</title>
            <style>
                body {{ font-family: Arial, sans-serif; max-width: 1000px; margin: 40px auto; padding: 0 20px; background: #ffffff; color: #111827; }}
                h1 {{ margin-bottom: 8px; }}
                p {{ color: #444; line-height: 1.5; }}
                .summary {{ margin-top: 22px; padding: 18px; border-radius: 14px; background: #f3f4f6; border: 1px solid #e5e7eb; }}
                .kpi {{ margin: 8px 0; font-size: 15px; }}
                .download-btn {{ display: inline-block; margin-top: 18px; background: #111827; color: white; text-decoration: none; padding: 12px 18px; border-radius: 8px; }}
                .download-btn:hover {{ opacity: 0.92; }}
                .cols {{ display: grid; grid-template-columns: 1fr 1fr; gap: 18px; margin-top: 20px; }}
                .card {{ background: #f9fafb; border: 1px solid #e5e7eb; padding: 18px; border-radius: 14px; }}
                .muted {{ color: #6b7280; }}
                textarea {{ width: 100%; height: 180px; font-size: 14px; padding: 12px; border: 1px solid #d1d5db; border-radius: 10px; }}
                button {{ background: #111827; color: white; border: none; padding: 12px 18px; font-size: 14px; border-radius: 8px; cursor: pointer; }}
                @media (max-width: 800px) {{
                    .cols {{ grid-template-columns: 1fr; }}
                }}
            </style>
        </head>
        <body>
            <h1>Resultado de consulta</h1>
            <p>La consulta terminó correctamente. Aquí puedes ver cuántas placas y registros se encontraron en cada hoja antes de descargar el Excel.</p>

            <div class="summary">
                <div class="kpi"><b>Placas consultadas:</b> {total_placas_consultadas}</div>
                <div class="kpi"><b>Placas encontradas en Escenarios:</b> {total_placas_esc}</div>
                <div class="kpi"><b>Placas encontradas en Escenarios_Estimado:</b> {total_placas_est}</div>
                <div class="kpi"><b>Registros encontrados en Escenarios:</b> {total_registros_esc}</div>
                <div class="kpi"><b>Registros encontrados en Escenarios_Estimado:</b> {total_registros_est}</div>

                <a class="download-btn" href="/descargar?placas={requests.utils.quote(placas_query)}">
                    Descargar Excel filtrado
                </a>
            </div>

            <div class="cols">
                <div class="card">
                    <h3>Placas encontradas en Escenarios</h3>
                    <p>{placas_esc_html}</p>
                </div>

                <div class="card">
                    <h3>Placas encontradas en Escenarios_Estimado</h3>
                    <p>{placas_est_html}</p>
                </div>
            </div>

            <div class="card" style="margin-top:20px;">
                <h3>Consultar de nuevo</h3>
                <form action="/consultar-form" method="post">
                    <textarea name="placas_texto">{html_escape(placas_query)}</textarea>
                    <br><br>
                    <button type="submit">Consultar placas</button>
                </form>
            </div>
        </body>
    </html>
    """