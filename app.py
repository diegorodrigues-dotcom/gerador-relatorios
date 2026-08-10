import streamlit as st
import docx
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
import io

# Configuração da página Web
st.set_page_config(page_title="Gerador de Relatórios Técnicos", layout="wide", page_icon="⚙️")

st.title("⚙️ Gerador Automático de Relatórios Técnicos")
st.subheader("Lavadoras de Alta Pressão e Aspiradores")

st.markdown("---")

# 1. CABEÇALHO / IDENTIFICAÇÃO
st.header("1. Informações Gerais do Ensaio")
col1, col2, col3 = st.columns(3)

with col1:
    codigo_st = st.text_input("Código ST / OS", value="ST 001942")
    tecnico = st.text_input("Técnico Responsável", value="Diego Rodrigues")
    data_ensaio = st.text_input("Data do Teste", value="01/06/2026")

with col2:
    item_testado = st.text_input("Item Testado", value="Olinda PW K5")
    quantidade = st.number_input("Quantidade de Amostras", min_value=1, value=4)
    modelo_motobomba = st.text_input("Modelo Motobomba", value="HY-603B")

with col3:
    objetivo = st.text_area("Objetivo do Teste", value="Analisar a funcionalidade e a durabilidade do equipamento", height=100)
    normas = st.text_area("Normas e Critérios", value="1- PFC (9.300-020.0)\n2- KN 082.023 cap. 4.7.1 / KN 082.021 cap.6", height=100)

conclusao_texto = st.text_area("Conclusão / Parecer Técnico Geral", 
    value="Amostras 1 e 2 (127V) e amostra 2 (220V) apresentaram defeitos semelhantes como água no cárter. Apesar das falhas identificadas, as amostras foram consideradas APROVADAS por atenderem a durabilidade mínima e não representarem risco ao usuário.", 
    height=120)

st.markdown("---")

# 2. DADOS DOS TESTES FUNCIONAIS (PARÂMETROS INICIAIS)
st.header("2. Parâmetros Iniciais (Teste Funcional)")

def criar_tabela_parametros(voltagem_label):
    st.subheader(f"Medições - {voltagem_label}")
    c1, c2, c3, c4, c5 = st.columns(5)
    
    with c1:
        st.write("**Tensão (V)**")
        v30 = st.number_input(f"30s ({voltagem_label})", value=126.7 if "127" in voltagem_label else 220.0, key=f"v30_{voltagem_label}")
        v1m = st.number_input(f"1min ({voltagem_label})", value=126.6 if "127" in voltagem_label else 220.0, key=f"v1m_{voltagem_label}")
        v5m = st.number_input(f"5min ({voltagem_label})", value=126.9 if "127" in voltagem_label else 220.0, key=f"v5m_{voltagem_label}")
    
    with c2:
        st.write("**Potência (kW)**")
        p30 = st.number_input(f"P 30s ({voltagem_label})", value=1.68, key=f"p30_{voltagem_label}")
        p1m = st.number_input(f"P 1min ({voltagem_label})", value=1.66, key=f"p1m_{voltagem_label}")
        p5m = st.number_input(f"P 5min ({voltagem_label})", value=1.64, key=f"p5m_{voltagem_label}")

    with c3:
        st.write("**Pressão (bar/PSI)**")
        pr30 = st.number_input(f"Pr 30s ({voltagem_label})", value=101.0, key=f"pr30_{voltagem_label}")
        pr1m = st.number_input(f"Pr 1min ({voltagem_label})", value=100.7, key=f"pr1m_{voltagem_label}")
        pr5m = st.number_input(f"Pr 5min ({voltagem_label})", value=101.1, key=f"pr5m_{voltagem_label}")

    with c4:
        st.write("**Vazão (l/h)**")
        vz30 = st.number_input(f"Vz 30s ({voltagem_label})", value=328.0, key=f"vz30_{voltagem_label}")
        vz1m = st.number_input(f"Vz 1min ({voltagem_label})", value=326.0, key=f"vz1m_{voltagem_label}")
        vz5m = st.number_input(f"Vz 5min ({voltagem_label})", value=327.0, key=f"vz5m_{voltagem_label}")

    with c5:
        st.write("**Corrente (A)**")
        i30 = st.number_input(f"I 30s ({voltagem_label})", value=13.9, key=f"i30_{voltagem_label}")
        i1m = st.number_input(f"I 1min ({voltagem_label})", value=13.7, key=f"i1m_{voltagem_label}")
        i5m = st.number_input(f"I 5min ({voltagem_label})", value=13.5, key=f"i5m_{voltagem_label}")

    # Cálculos automáticos das médias
    media_v = round((v30 + v1m + v5m) / 3, 2)
    media_p = round((p30 + p1m + p5m) / 3, 2)
    media_pr = round((pr30 + pr1m + pr5m) / 3, 2)
    media_vz = round((vz30 + vz1m + vz5m) / 3, 2)
    media_i = round((i30 + i1m + i5m) / 3, 2)

    st.info(f"📊 **Médias Calculadas ({voltagem_label}):** Tensão: {media_v}V | Potência: {media_p}kW | Pressão: {media_pr} | Vazão: {media_vz}l/h | Corrente: {media_i}A")
    
    return {
        "v30": v30, "v1m": v1m, "v5m": v5m, "med_v": media_v,
        "p30": p30, "p1m": p1m, "p5m": p5m, "med_p": media_p,
        "pr30": pr30, "pr1m": pr1m, "pr5m": pr5m, "med_pr": media_pr,
        "vz30": vz30, "vz1m": vz1m, "vz5m": vz5m, "med_vz": media_vz,
        "i30": i30, "i1m": i1m, "i5m": i5m, "med_i": media_i,
    }

dados_127v = criar_tabela_parametros("127V")
st.markdown("---")
dados_220v = criar_tabela_parametros("220V")

st.markdown("---")

# 3. UPLOAD DE FOTOS E CADASTRO DE FALHAS
st.header("3. Análise Visual por Amostra (Fotos e Defeitos)")

num_amostras = st.number_input("Quantas amostras deseja analisar com fotos?", min_value=1, max_value=10, value=2)

amostras_dados = []

for idx in range(num_amostras):
    st.subheader(f"Amostra #{idx+1}")
    col_a, col_b, col_c = st.columns([2, 2, 4])
    
    with col_a:
        nome_amostra = st.text_input(f"Identificação da Amostra #{idx+1}", value=f"Amostra {idx+1} 127V", key=f"nome_am_{idx}")
    with col_b:
        horas_ensaio = st.text_input(f"Horas de Ensaio", value="141 h", key=f"horas_am_{idx}")
        
    with col_c:
        defeitos_texto = st.text_area(f"Lista de Defeitos / Falhas Encontradas", 
                                      value="1- Parafuso de fixação solto;\n2- Cárter trincado;\n3- Água no cárter;\n4- Pistão travado;", 
                                      key=f"def_am_{idx}", height=100)
    
    fotos_uploaded = st.file_uploader(f"Anexar Fotos para {nome_amostra}", type=["png", "jpg", "jpeg"], accept_multiple_files=True, key=f"foto_am_{idx}")
    
    amostras_dados.append({
        "nome": nome_amostra,
        "horas": horas_ensaio,
        "defeitos": defeitos_texto,
        "fotos": fotos_uploaded
    })
    st.markdown("---")

# BOTÃO DE GERAÇÃO DO WORD
if st.button("🚀 GERAR RELATÓRIO EM WORD (.DOCX)", type="primary"):
    doc = docx.Document()

    # Título Principal
    p_title = doc.add_paragraph()
    run_title = p_title.add_run("RELATÓRIO TÉCNICO DE ENSAIO E QUALIDADE")
    run_title.bold = True
    run_title.font.size = Pt(16)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Tabela 1: Cabeçalho Geral
    table_hdr = doc.add_table(rows=6, cols=2)
    table_hdr.alignment = WD_TABLE_ALIGNMENT.CENTER
    table_hdr.style = 'Table Grid'

    dados_hdr = [
        ("Código ST:", codigo_st),
        ("Técnico Responsável:", tecnico),
        ("Data do Ensaio:", data_ensaio),
        ("Item / Equipamento Testado:", item_testado),
        ("Quantidade de Amostras:", str(quantidade)),
        ("Modelo da Motobomba:", modelo_motobomba)
    ]

    for i, (label, val) in enumerate(dados_hdr):
        row = table_hdr.rows[i]
        row.cells[0].paragraphs[0].add_run(label).bold = True
        row.cells[1].paragraphs[0].add_run(val)

    doc.add_paragraph()

    # Objetivos e Conclusão
    doc.add_heading("1. Objetivo do Teste", level=2)
    doc.add_paragraph(objetivo)

    doc.add_heading("2. Normas e Critérios de Aprovação", level=2)
    doc.add_paragraph(normas)

    doc.add_heading("3. Conclusão e Parecer Técnico", level=2)
    p_conc = doc.add_paragraph(conclusao_texto)

    # Seção de Fotos e Defeitos
    doc.add_heading("4. Análise Fotográfica e Inspeção das Amostras", level=2)

    for am in amostras_dados:
        doc.add_heading(f"• {am['nome']} ({am['horas']})", level=3)
        
        # Inserção das fotos
        if am["fotos"]:
            table_fotos = doc.add_table(rows=1, cols=min(len(am["fotos"]), 3))
            table_fotos.alignment = WD_TABLE_ALIGNMENT.CENTER
            
            for f_idx, foto_file in enumerate(am["fotos"]):
                if f_idx < 3: # Limita a 3 fotos por linha na tabela Word
                    img_bytes = foto_file.read()
                    image_stream = io.BytesIO(img_bytes)
                    
                    cell = table_fotos.rows[0].cells[f_idx]
                    p_img = cell.paragraphs[0]
                    p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p_img.add_run().add_picture(image_stream, width=Inches(1.8))
        
        doc.add_paragraph("Defeitos e Falhas Identificados:")
        doc.add_paragraph(am["defeitos"])
        doc.add_paragraph()

    # Salva o arquivo na memória para Download
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)

    st.success("✅ Relatório gerado com sucesso!")
    st.download_button(
        label="📥 Baixar Relatório Word (.docx)",
        data=buffer,
        file_name=f"{codigo_st} - {item_testado}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
