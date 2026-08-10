import streamlit as st
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn
import io

# Configuração da página Web
st.set_page_config(page_title="Gerador de Relatórios Técnicos", layout="wide", page_icon="⚙️")

st.title("⚙️ Gerador de Relatórios - Padrão Oficial")
st.subheader("Lavadoras de Alta Pressão e Aspiradores")

st.markdown("---")

# Funções auxiliares de formatação de tabelas do Word
def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for margin_name, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{margin_name}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

# 1. CABEÇALHO / IDENTIFICAÇÃO
st.header("1. Informações Gerais do Ensaio")
c1, c2, c3 = st.columns(3)

with c1:
    codigo_st = st.text_input("Código ST / OS", value="001942")
    tecnico = st.text_input("Técnico Responsável", value="Diego Rodrigues")
    data_ensaio = st.text_input("Data do Teste", value="01/06/2026")

with c2:
    item_testado = st.text_input("Item Testado", value="Olinda PW K5")
    quantidade = st.number_input("Quantidade de Amostras", min_value=1, value=4)
    modelo_motobomba = st.text_input("Modelo Motobomba", value="HY-603B")

with c3:
    objetivo = st.text_area("Objetivo do Teste", value="Analisar a funcionalidade e a durabilidade do equipamento", height=90)
    normas = st.text_area("Critério de Aprovação / Normas", value="1- Atender os critérios da PFC (9.300-020.0)\n2- Atender os critérios das normas KN 082.023 cap.4.7.1 / KN 082.021 cap.6", height=90)

conclusao_texto = st.text_area("Conclusão / Parecer Técnico Geral", 
    value="Modelo HY-603B\nAmostras 1 e 2 (127V) e amostra 2 (220V): Apresentaram defeitos semelhantes como água no cárter;\nAmostras 1 (127V) e amostra 2 (220V): Apresentaram a mesma falha, pistão travado;\nAmostras 1 e 2 (127V) e amostra 1 (220V): Tiveram defeitos como pistão riscado/oxidado;\nAmostras 2 (127V) e amostra 1 e 2 (220V): Tiveram defeitos como o parafuso de fixação da bomba com o motor solto;\nAmostra 1 e 2 (220V): Apresentou a mesma falha, o'ring do pistão do stop danificado;\nAmostra 1 e 2 (220V): Apresentou a mesma falha, corpo do Bypass rachado;\nEm conformidade com a norma KN 082.021, cap. 6 (Tabela Valor T), as amostras cujo tensão 127V demonstraram uma durabilidade superior à 115 horas, já as amostras cujo tensão 220V tendo uma média de 31 horas. Apesar das falhas identificadas, não representam riscos ao usuário, motivo pelo qual as amostras foram consideradas APROVADAS.", 
    height=140)

st.markdown("---")

# 2. TESTES FUNCIONAIS
st.header("2. Parâmetros Iniciais (Teste Funcional)")

def entrada_parametros(titulo, prefixo):
    st.subheader(titulo)
    col_a, col_b = st.columns(2)
    with col_a:
        partida = st.text_input(f"Partida a frio ({prefixo})", value="94V" if "127" in prefixo else "187V", key=f"part_{prefixo}")
    
    st.write("**Medições de Parâmetros (30s / 1min / 5min):**")
    m1, m2, m3, m4, m5 = st.columns(5)
    
    with m1:
        st.caption("Tensão (V)")
        v30 = st.number_input(f"V 30s", value=126.7 if "127" in prefixo else 221.8, key=f"v30_{prefixo}")
        v1m = st.number_input(f"V 1min", value=126.6 if "127" in prefixo else 220.9, key=f"v1m_{prefixo}")
        v5m = st.number_input(f"V 5min", value=126.9 if "127" in prefixo else 220.7, key=f"v5m_{prefixo}")

    with m2:
        st.caption("Potência (kW)")
        p30 = st.number_input(f"P 30s", value=1.688 if "127" in prefixo else 1.618, key=f"p30_{prefixo}")
        p1m = st.number_input(f"P 1min", value=1.665 if "127" in prefixo else 1.608, key=f"p1m_{prefixo}")
        p5m = st.number_input(f"P 5min", value=1.648 if "127" in prefixo else 1.575, key=f"p5m_{prefixo}")

    with m3:
        st.caption("Pressão bico")
        pr30 = st.number_input(f"Pr 30s", value=101.0 if "127" in prefixo else 104.4, key=f"pr30_{prefixo}")
        pr1m = st.number_input(f"Pr 1min", value=100.7 if "127" in prefixo else 104.4, key=f"pr1m_{prefixo}")
        pr5m = st.number_input(f"Pr 5min", value=101.1 if "127" in prefixo else 103.9, key=f"pr5m_{prefixo}")

    with m4:
        st.caption("Vazão (l/h)")
        vz30 = st.number_input(f"Vz 30s", value=328.0 if "127" in prefixo else 325.0, key=f"vz30_{prefixo}")
        vz1m = st.number_input(f"Vz 1min", value=326.0 if "127" in prefixo else 327.0, key=f"vz1m_{prefixo}")
        vz5m = st.number_input(f"Vz 5min", value=327.0 if "127" in prefixo else 327.0, key=f"vz5m_{prefixo}")

    with m5:
        st.caption("Corrente (A)")
        i30 = st.number_input(f"I 30s", value=13.98 if "127" in prefixo else 7.65, key=f"i30_{prefixo}")
        i1m = st.number_input(f"I 1min", value=13.75 if "127" in prefixo else 7.60, key=f"i1m_{prefixo}")
        i5m = st.number_input(f"I 5min", value=13.55 if "127" in prefixo else 7.48, key=f"i5m_{prefixo}")

    mv = round((v30 + v1m + v5m)/3, 1)
    mp = round((p30 + p1m + p5m)/3, 3)
    mpr = round((pr30 + pr1m + pr5m)/3, 1)
    mvz = round((vz30 + vz1m + vz5m)/3, 0)
    mi = round((i30 + i1m + i5m)/3, 2)

    return {
        "partida": partida,
        "v30": v30, "v1m": v1m, "v5m": v5m, "mv": mv,
        "p30": p30, "p1m": p1m, "p5m": p5m, "mp": mp,
        "pr30": pr30, "pr1m": pr1m, "pr5m": pr5m, "mpr": mpr,
        "vz30": vz30, "vz1m": vz1m, "vz5m": vz5m, "mvz": mvz,
        "i30": i30, "i1m": i1m, "i5m": i5m, "mi": mi
    }

dados_127v = entrada_parametros("1- Teste funcional Modelo HY-603B (127V Engate Rápido)", "127V")
st.markdown("---")
dados_220v = entrada_parametros("1.1- Teste funcional Modelo HY-603B (220V Engate Rápido)", "220V")

st.markdown("---")

# 3. FOTOS E DEFEITOS
st.header("3. Análise Visual e Galeria por Amostra")
num_amostras = st.number_input("Número de Amostras para Fotos", min_value=1, max_value=10, value=4)

amostras_dados = []
for idx in range(num_amostras):
    st.subheader(f"Amostra #{idx+1}")
    ca, cb, cc = st.columns([2, 2, 4])
    with ca:
        nome_am = st.text_input("Identificação", value=f"Amostra {idx+1} 127V", key=f"n_{idx}")
    with cb:
        horas_am = st.text_input("Duração / Horas", value="141 horas", key=f"h_{idx}")
    with cc:
        def_am = st.text_area("Lista de Defeitos", value="1- Faltando parafuso de fixação;\n2- Cárter trincado;\n3- Água no cárter;", key=f"d_{idx}", height=80)
    
    ft_am = st.file_uploader(f"Anexar Imagens para {nome_am}", type=["png", "jpg", "jpeg"], accept_multiple_files=True, key=f"f_{idx}")
    
    amostras_dados.append({"nome": nome_am, "horas": horas_am, "defeitos": def_am, "fotos": ft_am})
    st.markdown("---")

# BOTÃO DE GERAÇÃO DO WORD
if st.button("🚀 GERAR RELATÓRIO OFICIAL (.DOCX)", type="primary"):
    doc = docx.Document()

    # Ajuste de Margens
    sections = doc.sections
    for section in sections:
        section.top_margin = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin = Inches(0.8)
        section.right_margin = Inches(0.8)

    # Tabela 1: Cabeçalho com Estilo Identico
    tbl_top = doc.add_table(rows=1, cols=2)
    tbl_top.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl_top.autofit = False
    
    cell_st_label = tbl_top.rows[0].cells[0]
    cell_st_val = tbl_top.rows[0].cells[1]
    
    cell_st_label.width = Inches(1.5)
    cell_st_val.width = Inches(5.0)
    
    set_cell_background(cell_st_label, "1B365D") # Azul Escuro
    p0 = cell_st_label.paragraphs[0]
    r0 = p0.add_run("ST")
    r0.bold = True
    r0.font.color.rgb = RGBColor(255, 255, 255)
    r0.font.size = Pt(14)

    p1 = cell_st_val.paragraphs[0]
    r1 = p1.add_run(codigo_st)
    r1.bold = True
    r1.font.size = Pt(14)

    doc.add_paragraph()

    # Tabela de Dados Gerais
    tbl_meta = doc.add_table(rows=6, cols=2)
    tbl_meta.style = 'Table Grid'
    tbl_meta.alignment = WD_TABLE_ALIGNMENT.CENTER

    meta_info = [
        ("Teste realizado por", tecnico),
        ("Data", data_ensaio),
        ("Item testado", item_testado),
        ("Quantidade", str(quantidade)),
        ("Objetivo do teste", objetivo),
        ("Critério de aprovação", normas)
    ]

    for i, (k, v) in enumerate(meta_info):
        row = tbl_meta.rows[i]
        c0, c1 = row.cells[0], row.cells[1]
        c0.width = Inches(2.2)
        c1.width = Inches(4.3)
        
        set_cell_background(c0, "F2F2F2")
        p_k = c0.paragraphs[0]
        p_k.add_run(k).bold = True
        
        p_v = c1.paragraphs[0]
        p_v.add_run(v)

    doc.add_paragraph()

    # Conclusão
    p_conc_hdr = doc.add_paragraph()
    r_conc_h = p_conc_hdr.add_run("Conclusão")
    r_conc_h.bold = True
    r_conc_h.font.size = Pt(12)

    p_conc_b = doc.add_paragraph(conclusao_texto)
    doc.add_paragraph()

    # Tabelas Funcionais
    def adicionar_tabela_funcional(titulo, dados, sample_id):
        doc.add_paragraph().add_run(titulo).bold = True
        
        tbl = doc.add_table(rows=8, cols=8)
        tbl.style = 'Table Grid'
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER

        headers = ["Tempo de teste:", f"Partida a frio {dados['partida']}", "Tensão (V) - 60 Hz", "Potência absorvida (kW)", "Pressão com bico", "Vazão (l/h)", "Corrente (A)", "Sample ID"]
        
        hdr_row = tbl.rows[0]
        for col_i, h_text in enumerate(headers):
            cell = hdr_row.cells[col_i]
            set_cell_background(cell, "E6ECEF")
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.add_run(h_text).bold = True

        rows_data = [
            ("30s", "OK", str(dados["v30"]), str(dados["p30"]), str(dados["pr30"]), str(dados["vz30"]), str(dados["i30"]), sample_id),
            ("1 min", "", str(dados["v1m"]), str(dados["p1m"]), str(dados["pr1m"]), str(dados["vz1m"]), str(dados["i1m"]), ""),
            ("5 min", "", str(dados["v5m"]), str(dados["p5m"]), str(dados["pr5m"]), str(dados["vz5m"]), str(dados["i5m"]), ""),
            ("Média", "", str(dados["mv"]), str(dados["mp"]), str(dados["mpr"]), str(dados["mvz"]), str(dados["mi"]), "")
        ]

        for r_idx, r_vals in enumerate(rows_data):
            row = tbl.rows[r_idx + 1]
            for c_idx, val in enumerate(r_vals):
                cell = row.cells[c_idx]
                p = cell.paragraphs[0]
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                r = p.add_run(val)
                if r_vals[0] == "Média":
                    r.bold = True

    adicionar_tabela_funcional("1- Teste funcional Modelo " + modelo_motobomba, dados_127v, "AM 1")
    doc.add_paragraph()
    adicionar_tabela_funcional("1.1- Teste funcional Modelo " + modelo_motobomba, dados_220v, "AM 2")
    doc.add_paragraph()

    # Seção Visual de Fotos
    doc.add_paragraph().add_run("2- Durabilidade conforme norma KN 082.023 cap. 4.7.1").bold = True

    for am in amostras_dados:
        doc.add_paragraph().add_run(f"• {am['nome']}").bold = True
        
        # Galeria de fotos em grade
        if am["fotos"]:
            num_fotos = len(am["fotos"])
            cols_count = min(num_fotos, 3)
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

    # Salvar e disponibilizar download
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    st.success("✅ Relatório padrão oficial gerado com sucesso!")
    st.download_button(
        label="📥 Baixar Relatório Word (.docx)",
        data=buffer,
        file_name=f"ST {codigo_st} - {item_testado}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
