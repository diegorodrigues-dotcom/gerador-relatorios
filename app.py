import streamlit as st
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
import io

# Configuração da página Web
st.set_page_config(page_title="Gerador de Relatórios - Kärcher", layout="wide", page_icon="⚙️")

st.title("⚙️ Gerador de Relatórios Técnicos - Padrão Kärcher")
st.subheader("Lavadoras de Alta Pressão e Aspiradores")

st.markdown("---")

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

# 1. CABEÇALHO / IDENTIFICAÇÃO
st.header("1. Informações Gerais do Ensaio")
c1, c2, c3 = st.columns(3)

with c1:
    codigo_st = st.text_input("Código ST / OS", value="001942")
    tecnico = st.text_input("Técnico Responsável", value="Diego Rodrigues")
    data_ensaio = st.text_input("Data do Teste", value="01/06/2026")

with c2:
    item_testado = st.text_input("Item Testado", value="Olinda PW K5")
    modelo_motobomba = st.text_input("Modelo Motobomba", value="HY-603B")

with c3:
    objetivo = st.text_area("Objetivo do Teste", value="Analisar a funcionalidade e a durabilidade do equipamento", height=80)
    normas = st.text_area("Critério de Aprovação / Normas", value="1- PFC (9.300-020.0)\n2- KN 082.023 cap.4.7.1 / KN 082.021 cap.6", height=80)

conclusao_texto = st.text_area("Conclusão / Parecer Técnico Geral", 
    value="Amostras apresentaram boa performance nos testes funcionais e durabilidade dentro das especificações normativas.", 
    height=100)

st.markdown("---")

# 2. AMOSTRAS, PARÂMETROS E FOTOS
st.header("2. Cadastro de Amostras (Parâmetros + Fotos)")

num_amostras = st.number_input("Quantidade de Amostras para o Relatório", min_value=1, max_value=10, value=2)

amostras_dados = []

for idx in range(int(num_amostras)):
    st.markdown(f"### 🧪 Amostra #{idx+1}")
    
    col_a, col_b, col_c = st.columns(3)
    with col_a:
        sample_id = st.text_input(f"Identificação (Sample ID)", value=f"AM {idx+1}", key=f"id_{idx}")
        voltagem_conexao = st.text_input("Voltagem / Conexão", value="127V Engate Rápido" if idx==0 else "220V Engate Rápido", key=f"volt_{idx}")
    with col_b:
        partida = st.text_input("Partida a frio", value="94V" if "127" in voltagem_conexao else "187V", key=f"part_{idx}")
        horas_ensaio = st.text_input("Horas de Ensaio / Duração", value="141 horas", key=f"h_{idx}")
    with col_c:
        defeitos_texto = st.text_area("Lista de Defeitos / Falhas", value="1- Água no cárter;\n2- Pistão riscado;", key=f"def_{idx}", height=100)

    st.write(f"**Tabela de Medições da {sample_id} ({voltagem_conexao}):**")
    m1, m2, m3, m4, m5 = st.columns(5)
    
    with m1:
        st.caption("Tensão (V)")
        v30 = st.number_input("30s", value=126.7 if "127" in voltagem_conexao else 220.0, key=f"v30_{idx}")
        v1m = st.number_input("1min", value=126.6 if "127" in voltagem_conexao else 220.0, key=f"v1m_{idx}")
        v5m = st.number_input("5min", value=126.9 if "127" in voltagem_conexao else 220.0, key=f"v5m_{idx}")

    with m2:
        st.caption("Potência (kW)")
        p30 = st.number_input("30s", value=1.688, key=f"p30_{idx}")
        p1m = st.number_input("1min", value=1.665, key=f"p1m_{idx}")
        p5m = st.number_input("5min", value=1.648, key=f"p5m_{idx}")

    with m3:
        st.caption("Pressão bico")
        pr30 = st.number_input("30s", value=101.0, key=f"pr30_{idx}")
        pr1m = st.number_input("1min", value=100.7, key=f"pr1m_{idx}")
        pr5m = st.number_input("5min", value=101.1, key=f"pr5m_{idx}")

    with m4:
        st.caption("Vazão (l/h)")
        vz30 = st.number_input("30s", value=328.0, key=f"vz30_{idx}")
        vz1m = st.number_input("1min", value=326.0, key=f"vz1m_{idx}")
        vz5m = st.number_input("5min", value=327.0, key=f"vz5m_{idx}")

    with m5:
        st.caption("Corrente (A)")
        i30 = st.number_input("30s", value=13.98, key=f"i30_{idx}")
        i1m = st.number_input("1min", value=13.75, key=f"i1m_{idx}")
        i5m = st.number_input("5min", value=13.55, key=f"i5m_{idx}")

    mv = round((v30 + v1m + v5m)/3, 1)
    mp = round((p30 + p1m + p5m)/3, 3)
    mpr = round((pr30 + pr1m + pr5m)/3, 1)
    mvz = round((vz30 + vz1m + vz5m)/3, 0)
    mi = round((i30 + i1m + i5m)/3, 2)

    fotos_uploaded = st.file_uploader(f"Anexar Fotos para {sample_id}", type=["png", "jpg", "jpeg"], accept_multiple_files=True, key=f"foto_{idx}")

    amostras_dados.append({
        "sample_id": sample_id,
        "voltagem_conexao": voltagem_conexao,
        "partida": partida,
        "horas": horas_ensaio,
        "defeitos": defeitos_texto,
        "fotos": fotos_uploaded,
        "v30": v30, "v1m": v1m, "v5m": v5m, "mv": mv,
        "p30": p30, "p1m": p1m, "p5m": p5m, "mp": mp,
        "pr30": pr30, "pr1m": pr1m, "pr5m": pr5m, "mpr": mpr,
        "vz30": vz30, "vz1m": vz1m, "vz5m": vz5m, "mvz": mvz,
        "i30": i30, "i1m": i1m, "i5m": i5m, "mi": mi
    })
    st.markdown("---")

# BOTÃO DE GERAÇÃO DO WORD
if st.button("🚀 GERAR RELATÓRIO COMPLETO KÄRCHER (.DOCX)", type="primary"):
    doc = docx.Document()

    # Margens e Cabeçalho Kärcher
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)
        
        header = section.header
        p_hdr = header.paragraphs[0]
        p_hdr.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r_hdr = p_hdr.add_run("KÄRCHER - RELATÓRIO TÉCNICO DE ENSAIO")
        r_hdr.bold = True
        r_hdr.font.size = Pt(8)
        r_hdr.font.color.rgb = RGBColor(120, 120, 120)

    # Bloco ST
    tbl_top = doc.add_table(rows=1, cols=2)
    tbl_top.alignment = WD_TABLE_ALIGNMENT.CENTER
    c_st_label, c_st_val = tbl_top.rows[0].cells[0], tbl_top.rows[0].cells[1]
    c_st_label.width, c_st_val.width = Inches(1.5), Inches(5.0)
    
    set_cell_background(c_st_label, "FFED00")
    p0 = c_st_label.paragraphs[0]
    r0 = p0.add_run("ST")
    r0.bold, r0.font.size = True, Pt(14)

    p1 = c_st_val.paragraphs[0]
    r1 = p1.add_run(codigo_st)
    r1.bold, r1.font.size = True, Pt(14)

    doc.add_paragraph()

    # Tabela de Identificação Geral
    tbl_meta = doc.add_table(rows=6, cols=2)
    tbl_meta.style = 'Table Grid'
    tbl_meta.alignment = WD_TABLE_ALIGNMENT.CENTER

    meta_info = [
        ("Teste realizado por", tecnico),
        ("Data", data_ensaio),
        ("Item testado", item_testado),
        ("Quantidade", str(num_amostras)),
        ("Objetivo do teste", objetivo),
        ("Critério de aprovação", normas)
    ]

    for i, (k, v) in enumerate(meta_info):
        row = tbl_meta.rows[i]
        c0, c1 = row.cells[0], row.cells[1]
        c0.width, c1.width = Inches(2.2), Inches(4.3)
        set_cell_background(c0, "F2F2F2")
        c0.paragraphs[0].add_run(k).bold = True
        c1.paragraphs[0].add_run(v)

    doc.add_paragraph()

    # Conclusão
    doc.add_paragraph().add_run("Conclusão").bold = True
    doc.add_paragraph(conclusao_texto)
    doc.add_paragraph()

    # GERAR UMA TABELA DE PARÂMETROS PARA CADA AMOSTRA
    doc.add_paragraph().add_run("1- Teste Funcional de Parâmetros").bold = True

    for am in amostras_dados:
        p_sub = doc.add_paragraph()
        p_sub.add_run(f"Modelo {modelo_motobomba} - {am['sample_id']} ({am['voltagem_conexao']})").bold = True
        
        tbl = doc.add_table(rows=5, cols=8)
        tbl.style = 'Table Grid'
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER

        headers = ["Tempo de teste:", f"Partida a frio {am['partida']}", "Tensão (V) - 60 Hz", "Potência absorvida (kW)", "Pressão com bico", "Vazão (l/h)", "Corrente (A)", "Sample ID"]
        
        hdr_row = tbl.rows[0]
        for col_i, h_text in enumerate(headers):
            cell = hdr_row.cells[col_i]
            set_cell_background(cell, "FFED00")
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(h_text)
            r.bold = True
            r.font.size = Pt(9)

        rows_data = [
            ("30s", "OK", str(am["v30"]), str(am["p30"]), str(am["pr30"]), str(am["vz30"]), str(am["i30"]), am["sample_id"]),
            ("1 min", "", str(am["v1m"]), str(am["p1m"]), str(am["pr1m"]), str(am["vz1m"]), str(am["i1m"]), ""),
            ("5 min", "", str(am["v5m"]), str(am["p5m"]), str(am["pr5m"]), str(am["vz5m"]), str(am["i5m"]), ""),
            ("Média", "", str(am["mv"]), str(am["mp"]), str(am["mpr"]), str(am["mvz"]), str(am["mi"]), "")
        ]

        for r_idx, r_vals in enumerate(rows_data):
            row = tbl.rows[r_idx + 1]
            for c_idx, val in enumerate(r_vals):
                cell = row.cells[c_idx]
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r = p.add_run(val)
                r.font.size = Pt(9)
                if r_vals[0] == "Média":
                    r.bold = True

        doc.add_paragraph()

    # SEÇÃO DE FOTOS E DEFEITOS POR AMOSTRA
    doc.add_paragraph().add_run("2- Durabilidade e Análise Fotográfica").bold = True

    for am in amostras_dados:
        doc.add_paragraph().add_run(f"• {am['sample_id']} - {am['voltagem_conexao']} ({am['horas']})").bold = True
        
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

        doc.add_paragraph("Falhas / Defeitos identificados:")
        doc.add_paragraph(am["defeitos"])
        doc.add_paragraph()

    # Salvar e disponibilizar download
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    st.success("✅ Relatório gerado com tabelas individuais por amostra!")
    st.download_button(
        label="📥 Baixar Relatório Word (.docx)",
        data=buffer,
        file_name=f"ST {codigo_st} - Karcher {item_testado}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
