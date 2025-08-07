from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel
from typing import List, Dict, Any
from docxtpl import DocxTemplate
from tempfile import NamedTemporaryFile
from html2docx import html2docx
from io import BytesIO
import os
import uuid
import logging

# Configuração de logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI()

# CORS (pode restringir depois)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ===== MODELOS =====
class Item(BaseModel):
    nome: str
    valor: str

class DocumentoData(BaseModel):
    dados: List[Item]
    descricao: str  # HTML

# ===== UTILITÁRIOS =====
def carregar_template(caminho: str) -> DocxTemplate:
    if not os.path.exists(caminho):
        raise HTTPException(status_code=400, detail="Template não encontrado")
    return DocxTemplate(caminho)

def converter_html_para_subdoc(doc: DocxTemplate, html: str):
    try:
        logger.info("Convertendo HTML para subdocumento...")
        buffer: BytesIO = html2docx(html, title="Descrição")
    except Exception:
        raise HTTPException(status_code=400, detail="Erro ao converter HTML para DOCX")

    with NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
        tmp.write(buffer.getvalue())
        subdoc_path = tmp.name

    try:
        subdoc = doc.new_subdoc(subdoc_path)
        return subdoc
    except Exception:
        raise HTTPException(status_code=500, detail="Erro ao criar subdocumento")
    finally:
        os.unlink(subdoc_path)

def construir_contexto(data: DocumentoData, doc: DocxTemplate) -> Dict[str, Any]:
    contexto = {
        "dados": [{"nome": item.nome, "valor": item.valor} for item in data.dados],
        "descricao_subdoc": converter_html_para_subdoc(doc, data.descricao),
    }
    # Adição de novos campos:
    # contexto["outro_campo"] = data.outro_campo
    return contexto

def salvar_documento(doc: DocxTemplate) -> str:
    output_filename = f"/tmp/doc-{uuid.uuid4().hex[:8]}.docx"
    doc.save(output_filename)
    return output_filename

# ===== ENDPOINT =====
@app.post("/gerar-docx")
def gerar_docx(data: DocumentoData):
    try:
        logger.info("Recebendo requisição para gerar DOCX")
        template_path = "template.docx"

        doc = carregar_template(template_path)
        contexto = construir_contexto(data, doc)

        try:
            doc.render(contexto)
        except Exception as e:
            logger.error("Erro ao renderizar template", exc_info=True)
            raise HTTPException(status_code=500, detail="Erro ao renderizar documento")

        output_path = salvar_documento(doc)
        return FileResponse(
            output_path,
            filename="relatorio-tabela.docx",
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Erro inesperado: {str(e)}", exc_info=True)
        raise HTTPException(status_code=500, detail="Erro interno inesperado")
        
@app.get("/")
def root():
    return {"status": "ok", "message": "API está online"}
