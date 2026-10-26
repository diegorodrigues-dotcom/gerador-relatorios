import streamlit as st
import docx
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls
import io
import os
import statistics

# Configuração da página Web
st.set_page_config(page_title="Gerador de Relatórios - Kärcher", layout="wide", page_icon="⚙️")

st.title("⚙️ Gerador de Relatórios Técnicos - Padrão Kärcher")
st.subheader("Lavadora de Alta Pressão")

st.markdown("---")

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def prevent_row_split(row):
    """ Impede que o Word corte uma linha de tabela entre duas páginas """
    trPr = row._tr.get_or_add_trPr()
    trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))

def prevent_table_split(table):
    """ Impede que as linhas da tabela se dividam entre páginas """
    for row_idx, row in enumerate(table.rows):
        prevent_row_split(row)
        if row_idx < len(table.rows) - 1:
            for cell in row.cells:
                for p in cell.paragraphs:
                    p.paragraph_format.keep_with_next = True

def add_field(paragraph, field_type):
    run = paragraph.add_run()
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = field_type
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'separate')
    fldChar3 = OxmlElement('w:fldChar')
    fldChar3.set(qn('w:fldCharType'), 'end')
    
    r = run._r
    r.append(fldChar1)
    r.append(instrText)
    r.append(fldChar2)
    r.append(fldChar3)

def safe_float(val):
    """Converte texto digitado para float de forma segura para cálculos"""
    try:
        if not val:
            return None
        return float(str(val).replace(',', '.'))
    except ValueError:
        return None

def safe_stdev(lista_floats, decimais=2):
    """Calcula o desvio padrão de forma segura"""
    validos = [v for v in lista_floats if v is not None]
    if len(validos) > 1:
        sd = statistics.stdev(validos)
        return round(sd, decimais)
    return ""

def fmt_br(val):
    """Converte o float/string para o formato com vírgula (ex: 220,8)"""
    if val == "" or val is None:
        return ""
    return str(val).replace('.', ',')

def format_horas(val_texto):
    """ Adiciona automaticamente o sufixo 'horas' caso o utilizador digite apenas números """
    if not
