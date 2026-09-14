import streamlit as st
import docx
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import io
import os

# Configuração da página Web
st.set_page_config(page_title="Gerador de Relatórios - Kärcher", layout="wide", page_icon="⚙️")

st.title("⚙️ Gerador de Relatórios Técnicos - Padrão Kärcher")
st.subheader("Lavadoras de Alta Pressão e Equipamentos Motorizados")

st.markdown("---")

def set_cell_background(cell, fill_hex):
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

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
    """Converte texto digitado para float de forma segura para cálculo da média"""
    try:
        if not val:
            return 0.0
        return float(str(val).replace(',', '.'))
    except ValueError:
        return 0.0

# Verifica se a logo está presente na pasta
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

with c2:
    tecnico = st.text_input("Responsável", value="", placeholder="Digite seu nome...")
    normas = st.text_area("Critério de Aprovação / Normas", value="", height=80, placeholder="Digite os critérios / normas...")

with c3:
    item_testado = st.text_input("Item Testado", value="", placeholder="O que você está testando?...")
    data_ensaio = st.text_input("Data do Teste", value="", placeholder="Ex: 01/01/2026")

conclusao_texto = st.text_area("Conclusão / Parecer Técnico Geral", 
    value="", 
    height=120,
    placeholder="Digite aqui a conclusão e parecer técnico geral...")

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
        sample_id = st.text_input(f"Sample ID", value="", placeholder="Ex: AM 1", key=f"id_{idx}")
        voltagem_conexao = st.text_input("Tensão", value="", placeholder="127V / 220V", key=f"volt_{idx}")
    with col_b:
        col_p1, col_p2 = st.columns([1, 1])
        with col_p1:
            partiu_frio = st.selectbox("Partiu a frio?", ["Sim", "Não"], key=f"p_status_{idx}")
        with col_p2:
            tensao_partida = st.text_input("Tensão de Partida", value="", placeholder="Ex: 94V", key=f"p_tensao_{idx}")
        
        partida = f"{partiu_frio} ({tensao_partida})" if tensao_partida else partiu_frio
        horas_ensaio = st.text_input("Tempo de Teste / Horas", value="", placeholder="Ex: 116 h", key=f"h_{idx}")
    with col_c:
        defeitos_texto = st.text_area("Lista de Falhas", value="", placeholder="Digite as falhas encontradas...", key=f"def_{idx}", height=80)

    st.write(f"**Parâmetros de Teste Funcional - {sample_id if sample_id else f'Amostra {idx+1}'}:**")
    
    if incluir_rpm:
        m1, m2, m3, m4, m5, m6 = st.columns(6)
    else:
        m1, m2, m3, m4, m5 = st.columns(5)
    
    with m1:
        st.caption("Tensão (V)")
        v30 = st.text_input("30s (V)", value="", placeholder="Ex: 126.2", key=f"v30_{idx}")
        v1m = st.text_input("1min (V)", value="", placeholder="Ex: 127.5", key=f"v1m_{idx}")
        v3m = st.text_input("3min (V)", value="", placeholder="Ex: 128.0", key=f"v3m_{idx}")
        v5m = st.text_input("5min (V)", value="", placeholder="Ex: 127.1", key=f"v5m_{idx}")

    with m2:
        st.caption("Potência (kW)")
        p30 = st.text_input("30s (kW)", value="", placeholder="Ex: 1.53", key=f"p30_{idx}")
        p1m = st.text_input("1min (kW)", value="", placeholder="Ex: 1.54", key=f"p1m_{idx}")
        p3m = st.text_input("3min (kW)", value="", placeholder="Ex: 1.55", key=f"p3m_{idx}")
        p5m = st.text_input("5min (kW)", value="", placeholder="Ex: 1.50", key=f"p5m_{idx}")

    with m3:
        st.caption("Pressão bico")
        pr30 = st.text_input("30s (Bar)", value="", placeholder="Ex: 91.9", key=f"pr30_{idx}")
        pr1m = st.text_input("1min (Bar)", value="", placeholder="Ex: 93.2", key=f"pr1m_{idx}")
        pr3m = st.text_input("3min (Bar)", value="", placeholder="Ex: 94.0", key=f"pr3m_{idx}")
        pr5m = st.text_input("5min (Bar)", value="", placeholder="Ex: 94.5", key=f"pr5m_{idx}")

    with m4:
        st.caption("Vazão (l/h)")
        vz30 = st.text_input("30s (l/h)", value="", placeholder="Ex: 293", key=f"vz30_{idx}")
        vz1m = st.text_input("1min (l/h)", value="", placeholder="Ex: 295", key=f"vz1m_{idx}")
        vz3m = st.text_input("3min (l/h)", value="", placeholder="Ex: 296", key=f"vz3m_{idx}")
        vz5m = st.text_input("5min (l/h)", value="", placeholder="Ex: 297", key=f"vz5m_{idx}")

    with m5:
        st.caption("Corrente (A)")
        i30 = st.text_input("30s (A)", value="", placeholder="Ex: 12.68", key=f"i30_{idx}")
        i1m = st.text_input("1min (A)", value="", placeholder="Ex: 12.64", key=f"i1m_{idx}")
        i3m = st.text_input("3min (A)", value="", placeholder="Ex: 12.60", key=f"i3m_{idx}")
        i5m = st.text_input("5min (A)", value="", placeholder="Ex: 12.38", key=f"i5m_{idx}")

    rpm30, rpm1m, rpm3m, rpm5m, mrpm = "", "", "", "", ""
    if incluir_rpm:
        with m6:
            st.caption("RPM")
            rpm30 = st.text_input("30s (rpm)", value="", placeholder="Ex: 3450", key=f"rpm30_{idx}")
            rpm1m = st.text_input("1min (rpm)", value="", placeholder="Ex: 3435", key=f"rpm1m_{idx}")
            rpm3m = st.text_input("3min (rpm)", value="", placeholder="Ex: 3420", key=f"rpm3m_{idx}")
            rpm5m = st.text_input("5min (rpm)", value="", placeholder="Ex: 3410", key=f"rpm5m_{idx}")

    num_v = [safe_float(v30), safe_float(v1m), safe_float(v3m), safe_float(v5m)]
    num_p = [safe_float(p30), safe_float(p1m), safe_float(p3m), safe_float(p5m)]
    num_pr = [safe_float(pr30), safe_float(pr1m), safe_float(pr3m), safe_float(pr5m)]
    num_vz = [safe_float(vz30), safe_float(vz1m), safe_float(vz3m), safe_float(vz5m)]
    num_i = [safe_float(i30), safe_float(i1m), safe_float(i3m), safe_float(i5m)]

    mv = round(sum(num_v)/4, 1) if any(num_v) else ""
    mp = round(sum(num_p)/4, 2) if any(num_p) else ""
    mpr = round(sum(num_pr)/4, 1) if any(num_pr) else ""
    mvz = round(sum(num_vz)/4, 0) if any(num_vz) else ""
    mi = round(sum(num_i)/4, 2) if any(num_i) else ""

    if incluir_rpm:
        num_rpm = [safe_float(rpm30), safe_float(rpm1m), safe_float(rpm3m), safe_float(rpm5m)]
        calc_mrpm = round(sum(num_rpm)/4, 0) if any(num_rpm) else ""
        mrpm = str(int(calc_mrpm)) if calc_mrpm != "" else ""

    fotos_uploaded = st.file_uploader(f"Anexar Imagens para {sample_id if sample_id else f'Amostra {idx+1}'}", type=["png", "jpg", "jpeg"], accept_multiple_files=True, key=f"foto_{idx}")

    amostras_dados.append({
        "sample_id": sample_id,
        "voltagem_conexao": voltagem_conexao,
        "partida": partida,
        "horas": horas_ensaio,
        "defeitos": defeitos_texto,
        "fotos": fotos_uploaded,
        "v30": v30, "v1m": v1m, "v3m": v3m, "v5m": v5m, "mv": str(mv) if mv != "" else "",
        "p30": p30, "p1m": p1m, "p3m": p3m, "p5m": p5m, "mp": str(mp) if mp != "" else "",
        "pr30": pr30, "pr1m": pr1m, "pr3m": pr3m, "pr5m": pr5m, "mpr": str(mpr) if mpr != "" else "",
        "vz30": vz30, "vz1m": vz1m, "vz3m": vz3m, "vz5m": vz5m, "mvz": str(int(mvz)) if mvz != "" else "",
        "i30": i30, "i1m": i1m, "i3m": i3m, "i5m": i5m, "mi": str(mi) if mi != "" else "",
        "rpm30": rpm30, "rpm1m": rpm1m, "rpm3m": rpm3m, "rpm5m": rpm5m, "mrpm": mrpm
    })
    st.markdown("---")

# ----------------------------------------------------
# SEÇÃO DE GERAR RELATÓRIO WORD (.DOCX)
# ----------------------------------------------------
st.header("3. Opção de Download do Relatório")

if st.button("🚀 GERAR RELATÓRIO WORD (.DOCX)", type="primary", use_container_width=True):
    doc = docx.Document()

    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.header_distance = Inches(0.4)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)
        
        # ----------------------------------------------------
        # DEFININDO O CABEÇALHO OFICIAL DO WORD (HEADER)
        # ----------------------------------------------------
        header = section.header
        p_hdr_init = header.paragraphs[0]
        p_hdr_init.text = ""  # Limpa parágrafo padrão do cabeçalho
        
        tbl_hdr = header.add_table(rows=1, cols=2, width=Inches(6.9))
        tbl_hdr.autofit = False
        tbl_hdr.alignment = WD_TABLE_ALIGNMENT.CENTER
        c_left, c_right = tbl_hdr.rows[0].cells[0], tbl_hdr.rows[0].cells[1]
        
        # Largura da célula esquerda expandida para 12,0 cm para eliminar quebra de linha
        c_left.width = Cm(12.00)
        c_right.width = Cm(5.00)
        c_left.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        c_right.vertical_alignment = WD_ALIGN_VERTICAL.CENTER

        p_dept = c_left.paragraphs[0]
        p_dept.paragraph_format.space_before = Pt(0)
        p_dept.paragraph_format.space_after = Pt(0)
        r_dept = p_dept.add_run("Departamento de testes e desenvolvimentos")
        r_dept.font.name = 'Arial'
        r_dept.bold = True
        r_dept.font.size = Pt(12)

        p_logo = c_right.paragraphs[0]
        p_logo.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p_logo.paragraph_format.space_before = Pt(0)
        p_logo.paragraph_format.space_after = Pt(0)
        
        if os.path.exists(logo_path):
            p_logo.add_run().add_picture(logo_path, width=Cm(4.71), height=Cm(1.26))
        else:
            r_logo = p_logo.add_run("KÄRCHER")
            r_logo.font.name = 'Arial'
            r_logo.bold = True
            r_logo.font.size = Pt(18)
        
        # ----------------------------------------------------
        # DEFININDO O RODAPÉ OFICIAL DO WORD (FOOTER)
        # ----------------------------------------------------
        footer = section.footer
        p_ft = footer.paragraphs[0]
        p_ft.text = ""
        
        tbl_ft = footer.add_table(1, 3, Inches(6.9))
        tbl_ft.alignment = WD_TABLE_ALIGNMENT.CENTER
        
        c_st_ft, c_page_ft, c_item_ft = tbl_ft.rows[0].cells[0], tbl_ft.rows[0].cells[1], tbl_ft.rows[0].cells[2]
        c_st_ft.width, c_page_ft.width, c_item_ft.width = Inches(2.3), Inches(2.3), Inches(2.3)
        
        p_st_ft = c_st_ft.paragraphs[0]
        r_st = p_st_ft.add_run(f"ST {codigo_st}" if codigo_st else "ST")
        r_st.font.name = 'Arial'
        r_st.font.size = Pt(10)
        r_st.font.color.rgb = RGBColor(50, 50, 50)
        
        p_page = c_page_ft.paragraphs[0]
        p_page.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_p1 = p_page.add_run("Folha ")
        r_p1.font.name = 'Arial'
        r_p1.font.size = Pt(10)
        r_p1.font.color.rgb = RGBColor(50, 50, 50)
        add_field(p_page, 'PAGE')
        r_p2 = p_page.add_run(" de ")
        r_p2.font.name = 'Arial'
        r_p2.font.size = Pt(10)
        r_p2.font.color.rgb = RGBColor(50, 50, 50)
        add_field(p_page, 'NUMPAGES')
        
        p_item = c_item_ft.paragraphs[0]
        p_item.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r_it = p_item.add_run(item_testado if item_testado else "")
        r_it.font.name = 'Arial'
        r_it.font.size = Pt(10)
        r_it.font.color.rgb = RGBColor(50, 50, 50)

    # ----------------------------------------------------
    # CORPO DO DOCUMENTO
    # ----------------------------------------------------
    # TÍTULO: RELATÓRIO DE TESTE EM ARIAL TAMANHO 37
    p_main_title = doc.add_paragraph()
    p_main_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_main_title.paragraph_format.space_before = Pt(14)
    p_main_title.paragraph_format.space_after = Pt(20)
    r_title = p_main_title.add_run("Relatório de teste")
    r_title.font.name = 'Arial'
    r_title.bold = True
    r_title.font.size = Pt(37)

    # TABELA UNIFICADA DE INFORMAÇÕES GERAIS
    meta_info = [
        ("ST", codigo_st),
        ("Teste realizado por", tecnico),
        ("Data", data_ensaio),
        ("Item testado", item_testado),
        ("Quantidade", str(num_amostras)),
        ("Objetivo do teste", objetivo),
        ("Critério de aprovação", normas),
        ("Conclusão", conclusao_texto)
    ]

    tbl_meta = doc.add_table(rows=len(meta_info), cols=2)
    tbl_meta.autofit = False
    tbl_meta.style = 'Table Grid'
    tbl_meta.alignment = WD_TABLE_ALIGNMENT.CENTER

    for i, (k, v) in enumerate(meta_info):
        row = tbl_meta.rows[i]
        c0, c1 = row.cells[0], row.cells[1]
        
        c0.width = Cm(1.60)
        c1.width = Cm(13.50)
        
        set_cell_background(c0, "F2F2F2")
        
        r_k = c0.paragraphs[0].add_run(k)
        r_k.font.name = 'Arial'
        r_k.bold = True
        
        p_val = c1.paragraphs[0]
        linhas_v = v.split('\n') if v else [""]
        for idx_l, linha in enumerate(linhas_v):
            if idx_l > 0:
                p_val = c1.add_paragraph()
            r_v = p_val.add_run(linha)
            r_v.font.name = 'Arial'

    doc.add_paragraph()

    p_sec1 = doc.add_paragraph().add_run("1- Teste funcional de Parâmetros")
    p_sec1.font.name = 'Arial'
    p_sec1.bold = True

    for am in amostras_dados:
        p_sub = doc.add_paragraph()
        r_sub = p_sub.add_run(f"Máquina com conexão {am['voltagem_conexao']}")
        r_sub.font.name = 'Arial'
        r_sub.bold = True
        
        num_cols = 9 if incluir_rpm else 8
        tbl = doc.add_table(rows=6, cols=num_cols)
        tbl.style = 'Table Grid'
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER

        if incluir_rpm:
            headers = ["Tempo de teste:", f"Partida a frio {am['partida']}", "Tensão (V) - 60 Hz", "Potência absorvida (kW)", "Pressão com bico", "Vazão (l/h)", "Corrente (A)", "RPM (rpm)", "Sample ID"]
            rows_data = [
                ("30s", "OK", str(am["v30"]), str(am["p30"]), str(am["pr30"]), str(am["vz30"]), str(am["i30"]), str(am["rpm30"]), am["sample_id"]),
                ("1 min", "", str(am["v1m"]), str(am["p1m"]), str(am["pr1m"]), str(am["vz1m"]), str(am["i1m"]), str(am["rpm1m"]), ""),
                ("3 min", "", str(am["v3m"]), str(am["p3m"]), str(am["pr3m"]), str(am["vz3m"]), str(am["i3m"]), str(am["rpm3m"]), ""),
                ("5 min", "", str(am["v5m"]), str(am["p5m"]), str(am["pr5m"]), str(am["vz5m"]), str(am["i5m"]), str(am["rpm5m"]), ""),
                ("Média", "", str(am["mv"]), str(am["mp"]), str(am["mpr"]), str(am["mvz"]), str(am["mi"]), str(am["mrpm"]), "")
            ]
        else:
            headers = ["Tempo de teste:", f"Partida a frio {am['partida']}", "Tensão (V) - 60 Hz", "Potência absorvida (kW)", "Pressão com bico", "Vazão (l/h)", "Corrente (A)", "Sample ID"]
            rows_data = [
                ("30s", "OK", str(am["v30"]), str(am["p30"]), str(am["pr30"]), str(am["vz30"]), str(am["i30"]), am["sample_id"]),
                ("1 min", "", str(am["v1m"]), str(am["p1m"]), str(am["pr1m"]), str(am["vz1m"]), str(am["i1m"]), ""),
                ("3 min", "", str(am["v3m"]), str(am["p3m"]), str(am["pr3m"]), str(am["vz3m"]), str(am["i3m"]), ""),
                ("5 min", "", str(am["v5m"]), str(am["p5m"]), str(am["pr5m"]), str(am["vz5m"]), str(am["i5m"]), ""),
                ("Média", "", str(am["mv"]), str(am["mp"]), str(am["mpr"]), str(am["mvz"]), str(am["mi"]), "")
            ]
        
        hdr_row = tbl.rows[0]
        for col_i, h_text in enumerate(headers):
            cell = hdr_row.cells[col_i]
            set_cell_background(cell, "A6A6A6")
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(h_text)
            r.font.name = 'Arial'
            r.bold = True
            r.font.size = Pt(8.0 if incluir_rpm else 8.5)
            r.font.color.rgb = RGBColor(0, 0, 0)

        for r_idx, r_vals in enumerate(rows_data):
            row = tbl.rows[r_idx + 1]
            for c_idx, val in enumerate(r_vals):
                cell = row.cells[c_idx]
                
                if r_vals[0] == "Média" and c_idx in ([0] + list(range(2, num_cols-1))):
                    set_cell_background(cell, "A6A6A6")
                    
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r = p.add_run(val)
                r.font.name = 'Arial'
                r.font.size = Pt(8.5 if incluir_rpm else 9.0)
                if r_vals[0] == "Média":
                    r.bold = True

        doc.add_paragraph()

    p_sec2 = doc.add_paragraph().add_run("2- Durabilidade conforme norma KN 082.023 cap. 4.7.1")
    p_sec2.font.name = 'Arial'
    p_sec2.bold = True

    for am in amostras_dados:
        p_am = doc.add_paragraph()
        r_am = p_am.add_run(f"• {am['sample_id']} ({am['voltagem_conexao']})")
        r_am.font.name = 'Arial'
        r_am.bold = True
        
        if am["fotos"]:
            cols_count = min(len(am["fotos"]), 3)
            tbl_ft = doc.add_table(rows=1, cols=cols_count)
            tbl_ft.alignment = WD_TABLE_ALIGNMENT.CENTER
            
            for f_i, f_file in enumerate(am["fotos"]):
                if f_i < 3:
                    cell = tbl_ft.rows[0].cells[f_i]
                    p_img = cell.paragraphs[0]
                    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    img_stream = io.BytesIO(f_file.read())
                    p_img.add_run().add_picture(img_stream, width=Inches(1.8))

        p_fail_lbl = doc.add_paragraph(f"Após {am['horas']} foram identificadas as falhas / defeitos:")
        for r in p_fail_lbl.runs:
            r.font.name = 'Arial'
            
        p_fail_txt = doc.add_paragraph(am["defeitos"])
        for r in p_fail_txt.runs:
            r.font.name = 'Arial'
            
        doc.add_paragraph()

    p_res_lbl = doc.add_paragraph().add_run("Resumo das informações de defeitos e durabilidade apresentadas pelas amostras")
    p_res_lbl.font.name = 'Arial'
    p_res_lbl.bold = True
    
    tbl_res = doc.add_table(rows=len(amostras_dados)+1, cols=4)
    tbl_res.style = 'Table Grid'
    
    headers_res = ["Amostra", "Tempo de teste", "Vida útil esperada", "Falhas apresentadas"]
    for c_i, h_txt in enumerate(headers_res):
        cell = tbl_res.rows[0].cells[c_i]
        set_cell_background(cell, "A6A6A6")
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_h = p.add_run(h_txt)
        r_h.font.name = 'Arial'
        r_h.bold = True

    for r_i, am in enumerate(amostras_dados):
        row = tbl_res.rows[r_i+1]
        
        r_s = row.cells[0].paragraphs[0].add_run(am["sample_id"])
        r_s.font.name = 'Arial'
        
        r_h = row.cells[1].paragraphs[0].add_run(am["horas"])
        r_h.font.name = 'Arial'
        
        r_e = row.cells[2].paragraphs[0].add_run("60 h")
        r_e.font.name = 'Arial'
        
        r_f = row.cells[3].paragraphs[0].add_run("Identificadas no ensaio")
        r_f.font.name = 'Arial'

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    st.success("✅ Relatório Word (.docx) gerado com sucesso!")
    st.download_button(
        label="📥 Baixar Documento Word (.docx)",
        data=buffer,
        file_name=f"ST {codigo_st if codigo_st else '000'} - Karcher {item_testado if item_testado else 'Relatorio'}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True
    )
