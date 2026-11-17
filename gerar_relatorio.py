# Script de Automação de Relatórios Executivos - MVP
import os
import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

def gerar_grafico_redes():
    # Dados corporativos simulados por Rede (Insuficiência Financeira)
    redes = ['AGRO SP', 'SUL/MS/RO', 'SP/N/NO', 'AGRO A3', 'ESPECIAIS', 'AGRO SUL', 'MG/ES/RJ', 'AGRO A1']
    valores = [264.59, 236.23, 197.08, 145.09, 122.56, 74.30, 52.99, 5.98]
    
    fig, ax = plt.subplots(figsize=(7, 4.2))
    
    # Cores no padrão corporativo (Laranja/Vermelho Santander)
    barras = ax.barh(redes, valores, color='#E63946')
    ax.invert_yaxis()  # Inverte para o maior ficar no topo
    
    # Remove bordas desnecessárias
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_color('#CCCCCC')
    ax.spines['bottom'].set_color('#CCCCCC')
    
    ax.set_title('Insuficiência por Rede (Mi R$)', fontsize=12, fontweight='bold', color='#333333', pad=10)
    ax.tick_params(axis='x', labelsize=9)
    ax.tick_params(axis='y', labelsize=9)
    
    # Adiciona os rótulos de valor nas barras
    for barra in barras:
        width = barra.get_width()
        ax.text(width + 3, barra.get_y() + barra.get_height()/2, f'{width:.2f} Mi', 
                va='center', ha='left', fontsize=8, color='#333333', fontweight='bold')

    plt.tight_layout()
    caminho_img = 'grafico_rede_temp.png'
    plt.savefig(caminho_img, dpi=300, bbox_inches='tight')
    plt.close()
    return caminho_img

def criar_relatorio_executivo():
    # Define o diretório atual do script
    diretorio_atual = os.path.dirname(os.path.abspath(__file__))
    
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    
    blank_layout = prs.slide_layouts[6]
    segmentos = ["B1", "B2", "BD", "GIU", "Large", "Middle"]
    
    img_grafico = gerar_grafico_redes()
    
    for seg in segmentos:
        slide = prs.slides.add_slide(blank_layout)
        
        # --- CABEÇALHO ---
        title_box = slide.shapes.add_textbox(Inches(0.6), Inches(0.3), Inches(12.1), Inches(0.6))
        tf = title_box.text_frame
        p = tf.paragraphs[0]
        p.text = f"Insuficiência Financeira - Segmento {seg}"
        p.font.bold = True
        p.font.size = Pt(24)
        p.font.color.rgb = RGBColor(204, 0, 0) # Vermelho Santander
        
        # --- CARDS DE KPIS (TOPO) ---
        # Valor em Bi
        kpi1 = slide.shapes.add_textbox(Inches(0.6), Inches(1.0), Inches(3.5), Inches(0.8))
        kpi1.text_frame.paragraphs[0].text = "R$ 1.42 Bi\nValor Total em Insuficiência"
        kpi1.text_frame.paragraphs[0].font.bold = True
        kpi1.text_frame.paragraphs[0].font.size = Pt(13)
        
        # Qtd Contratos
        kpi2 = slide.shapes.add_textbox(Inches(4.5), Inches(1.0), Inches(3.5), Inches(0.8))
        kpi2.text_frame.paragraphs[0].text = "372\nQtde Contratos"
        kpi2.text_frame.paragraphs[0].font.bold = True
        kpi2.text_frame.paragraphs[0].font.size = Pt(13)
        
        # Qtd Clientes
        kpi3 = slide.shapes.add_textbox(Inches(8.4), Inches(1.0), Inches(3.5), Inches(0.8))
        kpi3.text_frame.paragraphs[0].text = "297\nQtde Clientes / CPFs"
        kpi3.text_frame.paragraphs[0].font.bold = True
        kpi3.text_frame.paragraphs[0].font.size = Pt(13)
        
        # --- GRÁFICO PRINCIPAL: INSUFICIÊNCIA POR REDE (LADO ESQUERDO) ---
        slide.shapes.add_picture(img_grafico, Inches(0.6), Inches(2.0), width=Inches(7.0))
        
        # --- BLOCOS AUXILIARES (LADO DIREITO) ---
        # Top 5 Garantias / Concentração
        box_top5 = slide.shapes.add_textbox(Inches(7.9), Inches(2.0), Inches(4.8), Inches(2.3))
        tf_t5 = box_top5.text_frame
        tf_t5.word_wrap = True
        p_t5 = tf_t5.paragraphs[0]
        p_t5.text = "Concentração - Top 5 Garantias"
        p_t5.font.bold = True
        p_t5.font.size = Pt(13)
        p_t5_sub = tf_t5.add_paragraph()
        p_t5_sub.text = "1. Imóveis Comerciais (42%)\n2. Aplicações Financeiras (25%)\n3. Veículos e Frotas (18%)\n4. Equipamentos Industriais (10%)\n5. Outros (5%)"
        p_t5_sub.font.size = Pt(11)
        
        # Faixa de Dias Insuficientes & Status (Regularizados / Pendentes)
        box_status = slide.shapes.add_textbox(Inches(7.9), Inches(4.6), Inches(4.8), Inches(2.3))
        tf_st = box_status.text_frame
        tf_st.word_wrap = True
        p_st = tf_st.paragraphs[0]
        p_st.text = "Status & Faixa de Dias (Cruzamento Temporal)"
        p_st.font.bold = True
        p_st.font.size = Pt(13)
        p_st_sub = tf_st.add_paragraph()
        p_st_sub.text = (
            "• Chave composta: Agência + Contrato + Garantia\n"
            "• Operações Regularizadas: Redução no ciclo\n"
            "• Faixas: 0-30 dias | 31-60 dias | Acima de 60 dias"
        )
        p_st_sub.font.size = Pt(11)

    # Salva na mesma pasta do script
    output_path = os.path.join(diretorio_atual, "Relatorio_Executivo_Garantias_Segmentos.pptx")
    prs.save(output_path)
    
    # Remove imagem temporária
    if os.path.exists(img_grafico):
        os.remove(img_grafico)
        
    print(f"[SUCESSO] Relatório gerado na mesma pasta: {output_path}")

if __name__ == "__main__":
    criar_relatorio_executivo()
