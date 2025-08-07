from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel
from typing import List
from docxtpl import DocxTemplate
from tempfile import NamedTemporaryFile
import uuid
import os
import logging
from html2docx import html2docx
from io import BytesIO

# Logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

app = FastAPI()

# CORS (ajustes de domínio)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Para produção, use: ["https://seusite.com"]
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Models
class Item(BaseModel):
    nome: str
    valor: float

class DocumentoData(BaseModel):
    dados: List[Item]
    descricao: str  # campo HTML

@app.post("/gerar-docx")
def gerar_docx(data: DocumentoData):
    try:
        logger.info("Iniciando geração de documento")

        # Verifica o template
        if not os.path.exists("template.docx"):
            raise HTTPException(status_code=400, detail="Template não encontrado")

        doc = DocxTemplate("template.docx")

        # Converte HTML para BytesIO
        try:
            logger.info(f"Convertendo HTML...")
            buf = html2docx(data.descricao, title="Descrição")
        except Exception as e:
            raise HTTPException(status_code=400, detail="HTML inválido")

        # Salva buffer em temp file
        with NamedTemporaryFile(delete=False, suffix='.docx') as tmp:
            tmp.write(buf.getvalue())
            subdoc_path = tmp.name

        # Cria subdocumento
        try:
            subdoc = doc.new_subdoc(subdoc_path)
        except Exception:
            raise HTTPException(status_code=500, detail="Erro ao criar subdocumento")
        finally:
            os.unlink(subdoc_path)

        # Contexto de substituição
        context = {
            "dados": [{"nome": item.nome, "valor": item.valor} for item in data.dados],
            "descricao_subdoc": subdoc
        }

        # Renderiza
        try:
            doc.render(context)
        except Exception:
            raise HTTPException(status_code=500, detail="Erro ao renderizar documento")

        # Salva resultado
        output_filename = f"/tmp/doc-tabela-{uuid.uuid4().hex[:8]}.docx"
        doc.save(output_filename)

        return FileResponse(
            output_filename,
            filename="relatorio-tabela.docx",
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except HTTPException:
        raise
    except Exception as e:
        logger.error(f"Erro inesperado: {str(e)}", exc_info=True)
        raise HTTPException(status_code=500, detail="Erro interno inesperado")
