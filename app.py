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
    if not val_texto:
        return ""
    val_clean = str(val_texto).strip()
    if val_clean.isdigit():
        return f"{val_clean} horas"
    if not val_clean.lower().endswith("horas") and not val_clean.lower().endswith("h"):
        return f"{val_clean} horas"
    if val_clean.lower().endswith("h"):
        val_num = val_clean[:-1].strip()
        if val_num.isdigit():
            return f"{val_num} horas"
    return val_clean

def format_voltagem(val_texto):
    """ Adiciona automaticamente o 'V' ao final da tensão se o utilizador não digitar """
    if not val_texto:
        return ""
    val_clean = str(val_texto).strip()
    if val_clean.isdigit():
        return f"{val_clean}V"
    if not val_clean.upper().endswith("V"):
        return f"{val_clean}V"
    return val_clean.upper()

# Verifica se a imagem do logótipo está presente no repositório
logo_path = os.path.join(os.path.dirname(__file__), "logo_karcher.png")
if not os.path.exists(logo_path):
    if os.path.exists("logo_karcher.png"):
        logo_path = "logo_karcher.png"
    else:
        st.warning("⚠️ Imagem 'logo_karcher.png' não foi encontrada no GitHub.")

# 1. CABEÇALHO / IDENTIFICAÇÃO GERAL
st.header("1. Informações Gerais do Ensaio")
c1, c2, c3 = st.columns(3)

with c1:
    codigo_st = st.text_input("ST", value="", placeholder="Ex: 001941")
    objetivo = st.text_area("Objetivo do Teste", value="", height=80, placeholder="Digite o objetivo do teste...")
    conclusao_texto = st.text_area("Conclusão / Parecer Técnico Geral", 
        value="", 
        height=120,
        placeholder="Digite aqui a conclusão e parecer técnico geral...")

with c2:
    tecnico = st.text_input("Responsável / Notificação (@ e-mail)", value="", placeholder="Ex: @nome.sobrenome@karcher.com")
    normas = st.text_area("Critério de Aprovação / Normas", value="", height=120, placeholder="Digite os critérios / normas...")

with c3:
    item_testado = st.text_input("Item Testado", value="", placeholder="O que você está testando?...")
    data_ensaio = st.text_input("Data do Teste", value="", placeholder="Ex: 01/01/2026")
    vida_util_input = st.text_input("Tempo de vida útil", value="", placeholder="Ex: 600")

# Formatação automática do tempo de vida útil
vida_util_geral = format_horas(vida_util_input)

st.markdown("---")

# 2. CADASTRO DINÂMICO DE AMOSTRAS E MEDIÇÕES
st.header("2. Cadastro de Amostras (Medições, Fotos e Defeitos)")

c_quant, c_rpm_opt = st.columns([1, 1])

with c_quant:
    num_amostras = st.number_input("Quantidade de Amostras", min_value=1, value=1)

with c_rpm_opt:
    incluir_rpm = st.selectbox("Equipamento possui medição de RPM?", ["Não", "Sim"]) == "Sim"

amostras_dados = []

for idx in range(int(num_amostras)):
    st.markdown(f"## Amostra {idx+1}")
    
    col_a, col_b, col_c = st.columns(3)
    with col_a:
        sample_id = st.text_input("Modelo", value="", placeholder="Ex: AM 1", key=f"id_{idx}")
        voltagem_input = st.text_input("Tensão (V)", value="", placeholder="127V / 220V", key=f"volt_{idx}")
        voltagem_conexao = format_voltagem(voltagem_input)
    with col_b:
        col_p1, col_p2 = st.columns([1, 1])
        with col_p1:
            partiu_frio = st.selectbox("Partiu a frio?", ["Sim", "Não"], key=f"p_status_{idx}")
        with col_p2:
            tensao_partida = st.text_input("Tensão de Partida", value="", placeholder="Ex: 94V", key=f"p_tensao_{idx}")
        
        partida_status = partiu_frio
        horas_input = st.text_input("Tempo de Teste / Horas", value="", placeholder="Ex: 116", key=f"h_{idx}")
        horas_ensaio = format_horas(horas_input)
    with col_c:
