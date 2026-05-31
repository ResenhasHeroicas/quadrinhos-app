from fastapi import FastAPI, HTTPException
from fastapi.staticfiles import StaticFiles
from fastapi.responses import HTMLResponse, StreamingResponse
from pydantic import BaseModel
from typing import Optional
import psycopg2
import psycopg2.extras
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import os
import io
import random
from datetime import date

app = FastAPI(title="Coleção de Quadrinhos")

DATABASE_URL = os.environ.get("DATABASE_URL", "")


def get_db():
    url = DATABASE_URL
    # Render usa "postgres://", psycopg2 precisa de "postgresql://"
    if url.startswith("postgres://"):
        url = url.replace("postgres://", "postgresql://", 1)
    conn = psycopg2.connect(url)
    return conn


def init_db():
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("""
                CREATE TABLE IF NOT EXISTS quadrinhos (
                    id SERIAL PRIMARY KEY,
                    titulo TEXT NOT NULL,
                    edicao TEXT,
                    editora TEXT,
                    categoria TEXT,
                    lido BOOLEAN DEFAULT FALSE,
                    obs TEXT,
                    inclusao TEXT,
                    created_at TIMESTAMP DEFAULT NOW()
                )
            """)
        conn.commit()


init_db()


class Quadrinho(BaseModel):
    titulo: str
    edicao: Optional[str] = None
    editora: Optional[str] = None
    categoria: Optional[str] = None
    lido: Optional[bool] = False
    obs: Optional[str] = None
    inclusao: Optional[str] = None


@app.get("/", response_class=HTMLResponse)
def index():
    with open("static/index.html", encoding="utf-8") as f:
        return f.read()


@app.get("/api/quadrinhos")
def listar(
    busca: str = "",
    categoria: str = "",
    editora: str = "",
    lido: str = "",
    page: int = 1,
    per_page: int = 50,
):
    where = ["1=1"]
    params = []

    if busca:
        where.append("(titulo ILIKE %s OR editora ILIKE %s OR obs ILIKE %s)")
        params += [f"%{busca}%", f"%{busca}%", f"%{busca}%"]
    if categoria:
        where.append("categoria = %s")
        params.append(categoria)
    if editora:
        where.append("editora = %s")
        params.append(editora)
    if lido == "sim":
        where.append("lido = TRUE")
    elif lido == "nao":
        where.append("lido = FALSE")

    where_sql = " AND ".join(where)
    offset = (page - 1) * per_page

    with get_db() as conn:
        with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
            cur.execute(f"SELECT COUNT(*) as total FROM quadrinhos WHERE {where_sql}", params)
            total = cur.fetchone()["total"]
            cur.execute(
                f"SELECT * FROM quadrinhos WHERE {where_sql} ORDER BY titulo, edicao LIMIT %s OFFSET %s",
                params + [per_page, offset],
            )
            items = cur.fetchall()

    return {"total": total, "page": page, "per_page": per_page, "items": [dict(r) for r in items]}


@app.get("/api/quadrinhos/{id}")
def obter(id: int):
    with get_db() as conn:
        with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
            cur.execute("SELECT * FROM quadrinhos WHERE id = %s", [id])
            row = cur.fetchone()
    if not row:
        raise HTTPException(404, "Não encontrado")
    return dict(row)


@app.post("/api/quadrinhos", status_code=201)
def criar(q: Quadrinho):
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute(
                "INSERT INTO quadrinhos (titulo, edicao, editora, categoria, lido, obs, inclusao) VALUES (%s,%s,%s,%s,%s,%s,%s) RETURNING id",
                [q.titulo, q.edicao, q.editora, q.categoria, q.lido, q.obs, q.inclusao or date.today().strftime("%d/%m/%Y")],
            )
            new_id = cur.fetchone()[0]
        conn.commit()
    return {"id": new_id}


@app.put("/api/quadrinhos/{id}")
def atualizar(id: int, q: Quadrinho):
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute(
                "UPDATE quadrinhos SET titulo=%s, edicao=%s, editora=%s, categoria=%s, lido=%s, obs=%s, inclusao=%s WHERE id=%s",
                [q.titulo, q.edicao, q.editora, q.categoria, q.lido, q.obs, q.inclusao, id],
            )
            if cur.rowcount == 0:
                raise HTTPException(404, "Não encontrado")
        conn.commit()
    return {"ok": True}


@app.delete("/api/quadrinhos/{id}")
def deletar(id: int):
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("DELETE FROM quadrinhos WHERE id = %s", [id])
            if cur.rowcount == 0:
                raise HTTPException(404, "Não encontrado")
        conn.commit()
    return {"ok": True}


@app.get("/api/stats")
def stats():
    with get_db() as conn:
        with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
            cur.execute("SELECT COUNT(*) as total, SUM(CASE WHEN lido THEN 1 ELSE 0 END) as lidos FROM quadrinhos")
            row = cur.fetchone()
    total = row["total"] or 0
    lidos = row["lidos"] or 0
    return {"total": total, "lidos": lidos, "nao_lidos": total - lidos}


@app.get("/api/filtros")
def filtros():
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT DISTINCT categoria FROM quadrinhos WHERE categoria IS NOT NULL ORDER BY categoria")
            categorias = [r[0] for r in cur.fetchall()]
            cur.execute("SELECT DISTINCT editora FROM quadrinhos WHERE editora IS NOT NULL ORDER BY editora")
            editoras = [r[0] for r in cur.fetchall()]
    return {"categorias": categorias, "editoras": editoras}


@app.get("/api/sugestoes")
def sugestoes():
    with get_db() as conn:
        with conn.cursor() as cur:
            cur.execute("SELECT DISTINCT titulo FROM quadrinhos WHERE titulo IS NOT NULL ORDER BY titulo")
            titulos = [r[0] for r in cur.fetchall()]
            cur.execute("SELECT DISTINCT editora FROM quadrinhos WHERE editora IS NOT NULL ORDER BY editora")
            editoras = [r[0] for r in cur.fetchall()]
            cur.execute("SELECT DISTINCT categoria FROM quadrinhos WHERE categoria IS NOT NULL ORDER BY categoria")
            categorias = [r[0] for r in cur.fetchall()]
    return {"titulos": titulos, "editoras": editoras, "categorias": categorias}


@app.get("/api/dashboard")
def dashboard(editora: str = "", categoria: str = "", status: str = ""):
    where = ["1=1"]
    params = []
    if editora and editora != "TODOS":
        where.append("editora = %s"); params.append(editora)
    if categoria and categoria != "TODOS":
        where.append("categoria = %s"); params.append(categoria)
    if status == "SIM":
        where.append("lido = TRUE")
    elif status == "NÃO":
        where.append("lido = FALSE")

    where_sql = " AND ".join(where)

    with get_db() as conn:
        with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
            cur.execute(f"SELECT COUNT(*) as total, SUM(CASE WHEN lido THEN 1 ELSE 0 END) as lidos FROM quadrinhos WHERE {where_sql}", params)
            row = cur.fetchone()
            total = row["total"] or 0
            lidos = row["lidos"] or 0

            cur.execute(f"SELECT categoria, COUNT(*) as qtd FROM quadrinhos WHERE {where_sql} AND categoria IS NOT NULL GROUP BY categoria ORDER BY qtd DESC", params)
            cat_rows = cur.fetchall()

            cur.execute(f"SELECT editora, COUNT(*) as qtd FROM quadrinhos WHERE {where_sql} AND editora IS NOT NULL GROUP BY editora ORDER BY qtd DESC", params)
            ed_rows = cur.fetchall()

            cur.execute(f"""
                SELECT TO_CHAR(TO_DATE(inclusao, 'DD/MM/YYYY'), 'YYYY-MM') as mes, COUNT(*) as qtd
                FROM quadrinhos WHERE {where_sql} AND inclusao IS NOT NULL AND inclusao ~ '^[0-9]{{2}}/[0-9]{{2}}/[0-9]{{4}}$'
                GROUP BY mes ORDER BY mes
            """, params)
            ev_rows = cur.fetchall()

    cat_count = {r["categoria"]: r["qtd"] for r in cat_rows}
    ed_sorted = [[r["editora"], r["qtd"]] for r in ed_rows]
    cat_sorted = [[r["categoria"], r["qtd"]] for r in cat_rows]
    evolucao = {r["mes"]: r["qtd"] for r in ev_rows if r["mes"]}

    return {
        "total": total,
        "lidos": lidos,
        "nao_lidos": total - lidos,
        "percentual": lidos / total if total else 0,
        "categorias": cat_count,
        "editoras": {r[0]: r[1] for r in ed_sorted},
        "evolucao": evolucao,
        "editoras_sorted": ed_sorted,
        "categorias_sorted": cat_sorted,
        "insights": {
            "top_editora": ed_sorted[0][0] if ed_sorted else "-",
            "top_categoria": cat_sorted[0][0] if cat_sorted else "-",
        },
    }


@app.get("/api/sugerir")
def sugerir():
    with get_db() as conn:
        with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
            cur.execute("SELECT titulo, edicao FROM quadrinhos WHERE lido = FALSE ORDER BY RANDOM() LIMIT 1")
            row = cur.fetchone()
    if not row:
        return {"titulo": "Você já leu tudo! 😎", "edicao": ""}
    return {"titulo": row["titulo"], "edicao": row["edicao"] or ""}


@app.get("/api/exportar")
def exportar():
    with get_db() as conn:
        with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
            cur.execute("SELECT titulo, edicao, editora, categoria, lido, obs, inclusao FROM quadrinhos ORDER BY titulo, edicao")
            rows = cur.fetchall()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "QUADRINHOS"

    header = ["TÍTULO", "EDIÇÃO", "EDITORA", "CATEGORIA", "LIDO", "OBS", "INCLUSAO"]
    widths  = [50, 12, 20, 12, 8, 25, 14]

    for col, (h, w) in enumerate(zip(header, widths), 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="4F46E5")
        cell.alignment = Alignment(horizontal="center")
        ws.column_dimensions[get_column_letter(col)].width = w

    for row_i, row in enumerate(rows, 2):
        ws.cell(row=row_i, column=1, value=row["titulo"])
        ws.cell(row=row_i, column=2, value=row["edicao"] or "")
        ws.cell(row=row_i, column=3, value=row["editora"] or "")
        ws.cell(row=row_i, column=4, value=row["categoria"] or "")
        ws.cell(row=row_i, column=5, value="SIM" if row["lido"] else "NÃO")
        ws.cell(row=row_i, column=6, value=row["obs"] or "")
        ws.cell(row=row_i, column=7, value=row["inclusao"] or "")

    ws.freeze_panes = "A2"

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)

    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": "attachment; filename=quadrinhos.xlsx"},
    )


@app.post("/api/importar-xlsx")
async def importar_xlsx():
    """Importa dados do arquivo quadrinhos.xlsx local (usado uma vez na migração)."""
    xlsx_path = "quadrinhos.xlsx"
    if not os.path.exists(xlsx_path):
        raise HTTPException(400, "Arquivo quadrinhos.xlsx não encontrado")

    wb = openpyxl.load_workbook(xlsx_path)
    ws = wb.active
    inserted = 0

    with get_db() as conn:
        with conn.cursor() as cur:
            for row in ws.iter_rows(min_row=2, values_only=True):
                titulo = str(row[0]).strip() if row[0] else ""
                if not titulo or titulo == "TÍTULO":
                    continue
                edicao   = str(row[1]).strip() if row[1] else None
                editora  = str(row[2]).strip() if row[2] else None
                categoria= str(row[3]).strip() if row[3] else None
                lido     = str(row[4]).strip().upper() == "SIM" if row[4] else False
                obs      = str(row[5]).strip() if row[5] else None
                inclusao = str(row[6]).strip() if row[6] else None
                cur.execute(
                    "INSERT INTO quadrinhos (titulo, edicao, editora, categoria, lido, obs, inclusao) VALUES (%s,%s,%s,%s,%s,%s,%s)",
                    [titulo, edicao, editora, categoria, lido, obs, inclusao]
                )
                inserted += 1
        conn.commit()

    return {"importados": inserted}


app.mount("/static", StaticFiles(directory="static"), name="static")
