import os
import re
import sys
import time
import pdfplumber
import asyncio
sys.path.append(os.path.join(os.path.dirname(__file__), '..'))
from excel.save_excel import save_data_to_excel

async def extract_data(pdf_path):
    extracted_data = {"empresa": None, "codigo": None, "total": None}
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    empresa_match = re.search(r"Empresa:\s*(.+?)(?:\s*Lote de Consolidação|$)", text)
                    codigo_match = re.search(r"Código Faturamento:\s*(\d+)", text)
                    total_match = re.search(r"Total Item:\s*([\d,\.]+)", text)

                    if empresa_match:
                        extracted_data["empresa"] = empresa_match.group(1).strip()
                    if codigo_match:
                        extracted_data["codigo"] = codigo_match.group(1).strip()
                    if total_match:
                        extracted_data["total"] = total_match.group(1).strip()
    except Exception as e:
        raise RuntimeError(f"Erro ao processar o PDF {pdf_path}: {e}")
    
    return extracted_data

async def process_pdf(pdf_path):
    return await extract_data(pdf_path)

async def process_pdfs(folder_path):
    start_time = time.time()
    all_data = []
    
    if not os.path.exists(folder_path):
        raise FileNotFoundError(f"A pasta {folder_path} não existe.")
    
    if not any(f.endswith('.pdf') for f in os.listdir(folder_path)):
        raise ValueError("Nenhum arquivo PDF encontrado na pasta.")
    
    tasks = []
    for filename in os.listdir(folder_path):
        if filename.endswith('.pdf'):
            pdf_path = os.path.join(folder_path, filename)
            tasks.append(process_pdf(pdf_path))
    
    try:
        all_data = await asyncio.gather(*tasks)
    except Exception as e:
        raise RuntimeError(f"Erro ao processar os PDFs: {str(e)}")
    
    parent_folder = os.path.dirname(folder_path)
    path = os.path.join(parent_folder, "resultado.xlsx")
    
    try:
        save_data_to_excel(all_data, path)
    except Exception as e:
        raise RuntimeError(f"Erro ao salvar a planilha: {str(e)}")

    end_time = time.time()
    execution_time = end_time - start_time
    print(f"Tempo de execução: {execution_time:.2f} segundos")

    return path