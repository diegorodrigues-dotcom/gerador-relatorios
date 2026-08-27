import streamlit as st
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import io

# Configuração da página Web
st.set_page_config(page_title="Gerador de Relatórios - Kärcher", layout="wide", page_icon="⚙️")

st.title("⚙️ Gerador de Relatórios Técnicos - Padrão Kärcher")
st.subheader("Lavadoras de Alta Pressão e Aspiradores")

st.markdown("---")

def set_cell_background(cell, fill_hex):
    from docx.oxml import parse_xml
    from docx.oxml.ns import nsdecls
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

# 1. CABEÇALHO / IDENTIFICAÇÃO GERAL
st.header("1. Informações Gerais do Ensaio")
c1, c2, c3 = st.columns(3)

with c1:
    codigo_st = st.text_input("Código ST", value="")
    tecnico = st.text_input("Técnico Responsável", value="")
    data_ensaio = st.text_input("Data do Teste", value="")

with c2:
    item_testado = st.text_input("Item Testado", value="")
    modelo_motobomba = st.text_input("Modelo Motobomba", value="")

with c3:
    objetivo = st.text_area("Objetivo do Teste", value="", height=80, placeholder="Digite o objetivo do teste...")
    normas = st.text_area("Critério de Aprovação / Normas", value="", height=80, placeholder="Digite as normas e critérios...")

conclusao_texto = st.text_area("Conclusão / Parecer Técnico Geral", 
    value="", 
    height=120,
    placeholder="Digite aqui a conclusão e parecer técnico geral...")

st.markdown("---")

# 2. CADASTRO DINÂMICO DE AMOSTRAS E MEDIÇÕES
st.header("2. Cadastro de Amostras (Medições, Fotos e Defeitos)")

num_amostras = st.number_input("Quantidade de Amostras", min_value=1, max_value=10, value=1)

amostras_dados = []

for idx in range(int(num_amostras)):
    st.markdown(f"### 🧪 Amostra #{idx+1}")
    
    col_a, col_b, col_c = st.columns(3)
    with col_a:
        sample_id = st.text_input(f"Sample ID", value="", placeholder="Ex: AM 1", key=f"id_{idx}")
        voltagem_conexao = st.text_input("Voltagem / Conexão", value="", placeholder="Ex: 127V Engate Rápido", key=f"volt_{idx}")
    with col_b:
        partida = st.text_input("Partida a frio", value="", placeholder="Ex: 94V", key=f"part_{idx}")
        horas_ensaio = st.text_input("Tempo de Teste / Horas", value="", placeholder="Ex: 116 h", key=f"h_{idx}")
    with col_c:
        defeitos_texto = st.text_area("Lista de Defeitos", value="", placeholder="1- Defeito A;\n2- Defeito B;", key=f"def_{idx}", height=100)

    st.write(f"**Parâmetros de Teste Funcional - {sample_id if sample_id else f'Amostra {idx+1}'}:**")
    m1, m2, m3, m4, m5 = st.columns(5)
    
    with m1:
        st.caption("Tensão (V)")
        v30 = st.number_input("30s (V)", value=0.0, step=0.1, key=f"v30_{idx}")
        v1m = st.number_input("1min (V)", value=0.0, step=0.1, key=f"v1m_{idx}")
        v5m = st.number_input("5min (V)", value=0.0, step=0.1, key=f"v5m_{idx}")

    with m2:
        st.caption("Potência (kW)")
        p30 = st.number_input("30s (kW)", value=0.00, step=0.01, format="%.2f", key=f"p30_{idx}")
        p1m = st.number_input("1min (kW)", value=0.00, step=0.01, format="%.2f", key=f"p1m_{idx}")
        p5m = st.number_input("5min (kW)", value=0.00, step=0.01, format="%.2f", key=f"p5m_{idx}")

    with m3:
        st.caption("Pressão bico")
        pr30 = st.number_input("30s (Pr)", value=0.0, step=0.1, key=f"pr30_{idx}")
        pr1m = st.number_input("1min (Pr)", value=0.0, step=0.1, key=f"pr1m_{idx}")
        pr5m = st.number_input("5min (Pr)", value=0.0, step=0.1, key=f"pr5m_{idx}")

    with m4:
        st.caption("Vazão (l/h)")
        vz30 = st.number_input("30s (Vz)", value=0.0, step=1.0, key=f"vz30_{idx}")
        vz1m = st.number_input("1min (Vz)", value=0.0, step=1.0, key=f"vz1m_{idx}")
        vz5m = st.number_input("5min (Vz)", value=0.0, step=1.0, key=f"vz5m_{idx}")

    with m5:
        st.caption("Corrente (A)")
        i30 = st.number_input("30s (A)", value=0.00, step=0.01, format="%.2f", key=f"i30_{idx}")
        i1m = st.number_input("1min (A)", value=0.00, step=0.01, format="%.2f", key=f"i1m_{idx}")
        i5m = st.number_input("5min (A)", value=0.00, step=0.01, format="%.2f", key=f"i5m_{idx}")

    mv = round((v30 + v1m + v5m)/3, 1)
    mp = round((p30 + p1m + p5m)/3, 2)
    mpr = round((pr30 + pr1m + pr5m)/3, 1)
    mvz = round((vz30 + vz1m + vz5m)/3, 0)
    mi = round((i30 + i1m + i5m)/3, 2)

    fotos_uploaded = st.file_uploader(f"Anexar Imagens para {sample_id if sample_id else f'Amostra {idx+1}'}", type=["png", "jpg", "jpeg"], accept_multiple_files=True, key=f"foto_{idx}")

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
if st.button("🚀 GERAR RELATÓRIO OFICIAL KÄRCHER (.DOCX)", type="primary"):
    doc = docx.Document()

    # Margens e Rodapé Kärcher
    for section in doc.sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)
        
        footer = section.footer
        p_ft = footer.paragraphs[0]
        p_ft.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        r_ft = p_ft.add_run("KÄRCHER - Relatório Técnico de Ensaio e Qualidade")
        r_ft.font.size = Pt(8)
        r_ft.font.color.rgb = RGBColor(120, 120, 120)

    # Bloco Limpo do Código ST
    tbl_top = doc.add_table(rows=1, cols=2)
    tbl_top.style = 'Table Grid'
    tbl_top.alignment = WD_TABLE_ALIGNMENT.CENTER
    c_st_label, c_st_val = tbl_top.rows[0].cells[0], tbl_top.rows[0].cells[1]
    c_st_label.width, c_st_val.width = Inches(1.2), Inches(5.3)
    
    p0 = c_st_label.paragraphs[0]
    r0 = p0.add_run("ST")
    r0.bold, r0.font.size = True, Pt(14)

    p1 = c_st_val.paragraphs[0]
    r1 = p1.add_run(codigo_st)
    r1.bold, r1.font.size = True, Pt(14)

    doc.add_paragraph()

    # Tabela de Dados Gerais
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

    # TABELAS DE PARÂMETROS CONFORME A QUANTIDADE DE AMOSTRAS
    doc.add_paragraph().add_run("1- Teste funcional Modelo " + modelo_motobomba).bold = True

    for am in amostras_dados:
        p_sub = doc.add_paragraph()
        p_sub.add_run(f"Máquina com conexão {am['voltagem_conexao']}").bold = True
        
        tbl = doc.add_table(rows=5, cols=8)
        tbl.style = 'Table Grid'
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER

        headers = ["Tempo de teste:", f"Partida a frio {am['partida']}", "Tensão (V) - 60 Hz", "Potência absorvida (kW)", "Pressão com bico", "Vazão (l/h)", "Corrente (A)", "Sample ID"]
        
        hdr_row = tbl.rows[0]
        for col_i, h_text in enumerate(headers):
            cell = hdr_row.cells[col_i]
            set_cell_background(cell, "FFED00") # Amarelo Kärcher no topo das tabelas
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(h_text)
            r.bold = True
            r.font.size = Pt(8.5)
            r.font.color.rgb = RGBColor(0, 0, 0)

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

    # ANÁLISE FOTOGRÁFICA
    doc.add_paragraph().add_run("2- Durabilidade conforme norma KN 082.023 cap. 4.7.1").bold = True

    for am in amostras_dados:
        doc.add_paragraph().add_run(f"• {am['sample_id']} ({am['voltagem_conexao']})").bold = True
        
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

        doc.add_paragraph(f"Após {am['horas']} foram identificadas as falhas / defeitos:")
        doc.add_paragraph(am["defeitos"])
        doc.add_paragraph()

    # MATRIZ FINAL DE DURABILIDADE
    doc.add_paragraph().add_run("Resumo das informações de defeitos e durabilidade apresentadas pelas amostras").bold = True
    
    tbl_res = doc.add_table(rows=len(amostras_dados)+1, cols=4)
    tbl_res.style = 'Table Grid'
    
    headers_res = ["Amostra", "Tempo de teste", "Vida útil esperada", "Falhas apresentadas"]
    for c_i, h_txt in enumerate(headers_res):
        cell = tbl_res.rows[0].cells[c_i]
        set_cell_background(cell, "FFED00")
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run(h_txt).bold = True

    for r_i, am in enumerate(amostras_dados):
        row = tbl_res.rows[r_i+1]
        row.cells[0].paragraphs[0].add_run(am["sample_id"])
        row.cells[1].paragraphs[0].add_run(am["horas"])
        row.cells[2].paragraphs[0].add_run("60 h")
        row.cells[3].paragraphs[0].add_run("Identificadas no ensaio")

    # Salvar e disponibilizar download
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    st.success("✅ Relatório padrão oficial Kärcher gerado com sucesso!")
    st.download_button(
        label="📥 Baixar Relatório Word (.docx)",
        data=buffer,
        file_name=f"ST {codigo_st if codigo_st else '000'} - Karcher {item_testado if item_testado else 'Relatorio'}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
