from fastapi import FastAPI, HTTPException, UploadFile, File
from fastapi.staticfiles import StaticFiles
from fastapi.responses import HTMLResponse
from pydantic import BaseModel
from typing import Optional
import sqlite3
import csv
import io
import os

app = FastAPI(title="Coleção de Quadrinhos")

DB_PATH = os.environ.get("DB_PATH", "quadrinhos.db")


def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    with get_db() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS quadrinhos (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                titulo TEXT NOT NULL,
                edicao TEXT,
                editora TEXT,
                categoria TEXT,
                lido INTEGER DEFAULT 0,
                obs TEXT,
                inclusao TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        """)


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
    with get_db() as conn:
        where = ["1=1"]
        params = []

        if busca:
            where.append("(titulo LIKE ? OR editora LIKE ? OR obs LIKE ?)")
            params += [f"%{busca}%", f"%{busca}%", f"%{busca}%"]
        if categoria:
            where.append("categoria = ?")
            params.append(categoria)
        if editora:
            where.append("editora = ?")
            params.append(editora)
        if lido != "":
            where.append("lido = ?")
            params.append(1 if lido == "sim" else 0)

        where_sql = " AND ".join(where)
        total = conn.execute(f"SELECT COUNT(*) FROM quadrinhos WHERE {where_sql}", params).fetchone()[0]
        offset = (page - 1) * per_page
        rows = conn.execute(
            f"SELECT * FROM quadrinhos WHERE {where_sql} ORDER BY titulo, edicao LIMIT ? OFFSET ?",
            params + [per_page, offset],
        ).fetchall()

        return {"total": total, "page": page, "per_page": per_page, "items": [dict(r) for r in rows]}


@app.get("/api/quadrinhos/{id}")
def obter(id: int):
    with get_db() as conn:
        row = conn.execute("SELECT * FROM quadrinhos WHERE id = ?", [id]).fetchone()
        if not row:
            raise HTTPException(404, "Não encontrado")
        return dict(row)


@app.post("/api/quadrinhos", status_code=201)
def criar(q: Quadrinho):
    with get_db() as conn:
        cur = conn.execute(
            "INSERT INTO quadrinhos (titulo, edicao, editora, categoria, lido, obs, inclusao) VALUES (?,?,?,?,?,?,?)",
            [q.titulo, q.edicao, q.editora, q.categoria, int(q.lido), q.obs, q.inclusao],
        )
        return {"id": cur.lastrowid}


@app.put("/api/quadrinhos/{id}")
def atualizar(id: int, q: Quadrinho):
    with get_db() as conn:
        res = conn.execute(
            "UPDATE quadrinhos SET titulo=?, edicao=?, editora=?, categoria=?, lido=?, obs=?, inclusao=? WHERE id=?",
            [q.titulo, q.edicao, q.editora, q.categoria, int(q.lido), q.obs, q.inclusao, id],
        )
        if res.rowcount == 0:
            raise HTTPException(404, "Não encontrado")
        return {"ok": True}


@app.delete("/api/quadrinhos/{id}")
def deletar(id: int):
    with get_db() as conn:
        res = conn.execute("DELETE FROM quadrinhos WHERE id = ?", [id])
        if res.rowcount == 0:
            raise HTTPException(404, "Não encontrado")
        return {"ok": True}


@app.get("/api/stats")
def stats():
    with get_db() as conn:
        total = conn.execute("SELECT COUNT(*) FROM quadrinhos").fetchone()[0]
        lidos = conn.execute("SELECT COUNT(*) FROM quadrinhos WHERE lido=1").fetchone()[0]
        por_categoria = conn.execute(
            "SELECT categoria, COUNT(*) as qtd FROM quadrinhos GROUP BY categoria ORDER BY qtd DESC"
        ).fetchall()
        por_editora = conn.execute(
            "SELECT editora, COUNT(*) as qtd FROM quadrinhos GROUP BY editora ORDER BY qtd DESC LIMIT 10"
        ).fetchall()
        return {
            "total": total,
            "lidos": lidos,
            "nao_lidos": total - lidos,
            "por_categoria": [dict(r) for r in por_categoria],
            "por_editora": [dict(r) for r in por_editora],
        }


@app.get("/api/filtros")
def filtros():
    with get_db() as conn:
        categorias = [r[0] for r in conn.execute("SELECT DISTINCT categoria FROM quadrinhos WHERE categoria IS NOT NULL ORDER BY categoria").fetchall()]
        editoras = [r[0] for r in conn.execute("SELECT DISTINCT editora FROM quadrinhos WHERE editora IS NOT NULL ORDER BY editora").fetchall()]
        return {"categorias": categorias, "editoras": editoras}


@app.post("/api/importar")
async def importar(file: UploadFile = File(...)):
    content = await file.read()
    text = content.decode("utf-8-sig")
    reader = csv.DictReader(io.StringIO(text))

    inserted = 0
    with get_db() as conn:
        for row in reader:
            titulo = row.get("TÍTULO") or row.get("TITULO") or row.get("título") or row.get("titulo") or ""
            if not titulo.strip():
                continue
            lido_raw = (row.get("LIDO") or "").strip().upper()
            lido = 1 if lido_raw == "SIM" else 0
            conn.execute(
                "INSERT INTO quadrinhos (titulo, edicao, editora, categoria, lido, obs, inclusao) VALUES (?,?,?,?,?,?,?)",
                [
                    titulo.strip(),
                    (row.get("EDIÇÃO") or row.get("EDICAO") or "").strip() or None,
                    (row.get("EDITORA") or "").strip() or None,
                    (row.get("CATEGORIA") or "").strip() or None,
                    lido,
                    (row.get("OBS") or "").strip() or None,
                    (row.get("INCLUSAO") or row.get("INCLUSÃO") or "").strip() or None,
                ],
            )
            inserted += 1

    return {"importados": inserted}


app.mount("/static", StaticFiles(directory="static"), name="static")
