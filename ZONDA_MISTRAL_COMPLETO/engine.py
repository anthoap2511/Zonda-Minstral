# -*- coding: utf-8 -*-
"""Motor ZONDA-MISTRAL: lectura de comprobantes y generación de Excel."""
from pathlib import Path
from datetime import datetime
import re
from copy import copy
import fitz
import openpyxl
from openpyxl.drawing.image import Image as ExcelImage

TEMPLATE = Path(__file__).resolve().parent / "engine" / "RENDICION DE CUENTAS N°8 OSCAR.xlsx"
SHEET_NAME = "RENDICION"
FIRST_RECEIPT_ROW, LAST_RECEIPT_ROW, MOBILITY_ROW = 11, 27, 28
COL = {"fecha":1,"ruc":2,"tipo":3,"serie":4,"numero":5,"proveedor":6,"detalle":7,"total":8}

def read_pdf(path):
    with fitz.open(path) as doc:
        return "\n".join(page.get_text("text") for page in doc)

def find(patterns, text):
    for pattern in patterns:
        m=re.search(pattern,text,re.I|re.M)
        if m: return m.group(1).strip()
    return None

def parse_receipt(path):
    text=read_pdf(path)
    lines=[x.strip() for x in text.splitlines() if x.strip()]
    rucs=re.findall(r"\b(?:R\.?\s*U\.?\s*C\.?|RUC)\s*[:\s]*([0-9]{11})\b",text,re.I)
    ruc=rucs[0] if rucs else None
    match=re.search(r"\b([FB]\d{3})[-\s]+(\d{1,8})\b",text,re.I)
    serie,numero=(match.group(1).upper(),match.group(2).zfill(8)) if match else (None,None)
    fecha_txt=find([
        r"FECHA\s+DE\s+EMISI[ÓO]N\s*[:\-]?\s*(\d{4}-\d{2}-\d{2})",
        r"FECHA\s+DE\s+EMISI[ÓO]N\s*[:\-]?\s*(\d{2}/\d{2}/\d{4})"],text)
    fecha=None
    if fecha_txt:
        for fmt in ("%Y-%m-%d","%d/%m/%Y"):
            try: fecha=datetime.strptime(fecha_txt,fmt); break
            except ValueError: pass
    proveedor=None
    for line in lines[:8]:
        up=line.upper()
        if any(x in up for x in ("S.A.C","S.A.","S.R.L","E.I.R.L")):
            proveedor=line; break
    if not proveedor and lines: proveedor=lines[0]
    total_txt=find([
        r"TOTAL\s+A\s+PAGAR\s*:\s*S/?\.?\s*([0-9,]+\.\d{2})",
        r"IMPORTE\s+TOTAL\s*:\s*S/?\.?\s*([0-9,]+\.\d{2})",
        r"TOTAL\s*:\s*S/?\.?\s*([0-9,]+\.\d{2})"],text)
    total=float(total_txt.replace(",","")) if total_txt else None
    descripcion=None
    for i,line in enumerate(lines):
        if line.upper() in ("NIU","UND","UNID","UNIDAD"):
            if i+1<len(lines) and not re.fullmatch(r"[\d.,]+",lines[i+1]):
                descripcion=lines[i+1]; break
    if not descripcion:
        descripcion=find([r"\b\d+\s+(?:NIU|UND|UNID)\s+(.+?)\s{2,}[0-9,]+\.\d{2}"],text)
    descripcion=descripcion or "REVISAR DESCRIPCIÓN"
    tipo="01" if serie and serie.startswith("F") else ("03" if serie and serie.startswith("B") else "")
    desc=descripcion.upper()
    if any(k in desc for k in ("HOTEL","HOSTAL","HOSPEDAJE","HABITACIÓN")): categoria="HOSPEDAJE"
    elif any(k in desc for k in ("TAXI","PASAJE","TRANSPORTE","BUS")): categoria="MOVILIDAD"
    elif any(k in desc for k in ("ALMUERZO","DESAYUNO","CENA","MENÚ","MENU","RESTAURANT")): categoria="ALIMENTO"
    else: categoria="OTROS"
    missing=[k for k,v in {"RUC":ruc,"serie/número":serie and numero,"fecha":fecha,"proveedor":proveedor,"total":total}.items() if not v]
    if missing: raise ValueError(f"{Path(path).name}: no se identificó {', '.join(missing)}")
    return {"fecha":fecha.strftime("%Y-%m-%d"),"ruc":ruc,"tipo":tipo,"serie":serie,"numero":numero,
            "proveedor":proveedor,"detalle":f"{categoria} - {descripcion}","total":total,"archivo":Path(path).name}

def copy_style(source,target):
    if source.has_style: target._style=copy(source._style)
    if source.number_format: target.number_format=source.number_format
    if source.alignment: target.alignment=copy(source.alignment)

def format_date(value):
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%d%m%Y"):
        try: return datetime.strptime(value,fmt)
        except ValueError: pass
    raise ValueError("La fecha de transferencia no tiene un formato válido.")

def make_excel(records, settings, output_path):
    if not TEMPLATE.exists(): raise FileNotFoundError(f"No se encuentra la plantilla: {TEMPLATE}")
    wb=openpyxl.load_workbook(TEMPLATE)
    if SHEET_NAME not in wb.sheetnames: raise ValueError(f"No existe la hoja {SHEET_NAME}.")
    ws=wb[SHEET_NAME]
    # Sustituye el logotipo anterior por el logotipo MISTRAL en la esquina superior izquierda.
    ws._images=[]
    logo_path=Path(__file__).resolve().parent / "engine" / "logo_mistral.png"
    if logo_path.exists():
        logo=ExcelImage(str(logo_path))
        logo.width, logo.height = 127, 47
        ws.add_image(logo, "A1")
    capacity=LAST_RECEIPT_ROW-FIRST_RECEIPT_ROW+1
    if len(records)>capacity:
        extra=len(records)-capacity
        ws.insert_rows(MOBILITY_ROW,amount=extra)
        for row in range(MOBILITY_ROW,MOBILITY_ROW+extra):
            for col in range(1,9): copy_style(ws.cell(LAST_RECEIPT_ROW,col),ws.cell(row,col))
        mobility_row=MOBILITY_ROW+extra
    else: mobility_row=MOBILITY_ROW
    for row in range(FIRST_RECEIPT_ROW,mobility_row):
        for col in range(1,9): ws.cell(row,col).value=None
    for idx,rec in enumerate(records):
        row=FIRST_RECEIPT_ROW+idx
        fecha=datetime.strptime(rec["fecha"],"%Y-%m-%d")
        ws.cell(row,1).value=fecha; ws.cell(row,1).number_format="d-mmm-yy"
        for col,key in [(2,"ruc"),(3,"tipo"),(4,"serie"),(5,"numero"),(6,"proveedor"),(7,"detalle"),(8,"total")]:
            ws.cell(row,col).value=rec[key]
        ws.cell(row,2).value=str(rec["ruc"]); ws.cell(row,5).value=str(rec["numero"])
        ws.cell(row,8).number_format='"S/" #,##0.00'
    total_row,received_row,balance_row=mobility_row+1,mobility_row+2,mobility_row+3
    ws.cell(total_row,8).value=f"=SUM(H{FIRST_RECEIPT_ROW}:H{mobility_row-1})"
    ws.cell(received_row,8).value="=D8"
    ws.cell(balance_row,8).value=f"=H{received_row}-H{total_row}"
    ws["A1"]=f"RENDICION DE CUENTAS N° {settings['numero_rendicion']}"
    ws["C3"]=settings["descripcion_objeto"]; ws["C4"]=settings["ubicacion"]
    ws["A8"]=format_date(settings["fecha_transferencia"]); ws["A8"].number_format="dd/mm/yy"
    ws["B8"]=str(settings["numero_cuenta"]); ws["D8"]=float(settings["importe_recibido"])
    ws["D8"].number_format='"S/"#,##0.00'
    # Texto de rendición con la estructura solicitada.
    ubicacion=(settings.get("ubicacion") or "").strip()
    # Si el campo contiene además la referencia de ciudad/mes/año, conservar
    # únicamente el destino (por ejemplo: "INDEPENDENCIA LIMA").
    ubicacion=re.sub(
        r"\\s+Y\\s+LIMA(?:\\s+(?:ENERO|FEBRERO|MARZO|ABRIL|MAYO|JUNIO|JULIO|AGOSTO|SEPTIEMBRE|OCTUBRE|NOVIEMBRE|DICIEMBRE)\\s+\\d{4})?\\s*$",
        "", ubicacion, flags=re.IGNORECASE
    ).strip()
    motivo=(settings.get("descripcion_objeto") or "").strip().rstrip(".")
    texto=(f"Yo, {settings['tecnico_nombre']}, identificado con DNI N° {settings['tecnico_dni']}, "
           f"rindo cuentas de los gastos de alimentacion, movilidad y hospedaje realizado desde Lima hacia "
           f"{ubicacion} con motivo de {motivo.lower()} según contrato con RENIEC.")
    ws["A9"]=texto
    ws.cell(mobility_row,1).value=(f"PLANILLA DE MOVILIDAD {settings['tecnico_nombre'].upper()}"
                                  if settings.get("incluir_movilidad", True) else "PLANILLA DE MOVILIDAD")
    row_shift=max(0,mobility_row-MOBILITY_ROW)
    ws.cell(34+row_shift,6).value=settings["fecha_firma"]
    ws.cell(40+row_shift,7).value=settings["tecnico_nombre"]
    ws.cell(41+row_shift,7).value=f"DNI N° {settings['tecnico_dni']}"
    Path(output_path).parent.mkdir(parents=True,exist_ok=True)
    wb.save(output_path)
    return output_path

def process_pdfs(pdf_paths, settings, output_path):
    if not pdf_paths: raise ValueError("Esta rendición no tiene comprobantes PDF.")
    records=[]; errors=[]
    for path in pdf_paths:
        try: records.append(parse_receipt(path))
        except Exception as exc: errors.append(str(exc))
    if errors: raise ValueError("No se generó el Excel. Revisa estos comprobantes:\\n- "+"\\n- ".join(errors))
    detalle_gasto=(settings.get("detalle_gasto") or "").strip()
    if detalle_gasto:
        for record in records:
            record["categoria_rendicion"]=record["detalle"].split(" - ")[0]
            record["detalle"]=detalle_gasto
    make_excel(records,settings,output_path)
    return records
