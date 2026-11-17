import pandas as pd
import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.util import Inches, Pt

def gerar_relatorio_com_grafico():
    print("Iniciando o processamento e geração de gráficos da Loja Hawara...")
    
    # 1. Base de dados simulada de vendas e estoque
    dados = {
        'Categoria': ['Camisetas', 'Calças', 'Jaquetas', 'Vestidos', 'Acessórios'],
        'Estoque_Atual': [150, 80, 45, 60, 200],
        'Vendas_Mes': [320, 140, 90, 110, 410]
    }
    
    df = pd.DataFrame(dados)
    
    # 2. Gerar gráfico profissional com Matplotlib
    plt.figure(figsize=(7, 4.5))
    x = range(len(df['Categoria']))
    largura = 0.35
    
    plt.bar([p - largura/2 for p in x], df['Vendas_Mes'], largura, label='Vendas do Mês', color='#1f77b4')
    plt.bar([p + largura/2 for p in x], df['Estoque_Atual'], largura, label='Estoque Atual', color='#ff7f0e')
    
    plt.xlabel('Categorias de Vestuário', fontweight='bold')
    plt.ylabel('Unidades', fontweight='bold')
    plt.title('Desempenho Operacional: Vendas vs. Estoque - Loja Hawara', fontweight='bold', fontsize=12)
    plt.xticks(x, df['Categoria'])
    plt.legend()
    plt.tight_layout()
    
    # Salva o gráfico como imagem temporária
    caminho_grafico = 'grafico_desempenho.png'
    plt.savefig(caminho_grafico, dpi=300)
    plt.close()
    print("Gráfico gerado e salvo com sucesso.")
    
    # 3. Criar a apresentação em PowerPoint e inserir o gráfico
    prs = Presentation()
    slide_layout = prs.slide_layouts[6] # Layout em branco para total liberdade visual
    slide = prs.slides.add_slide(slide_layout)
    
    # Título principal no topo do slide (utilizando Pt(24) corretamente)
    txBox = slide.shapes.add_textbox(Inches(0.8), Inches(0.5), Inches(8.5), Inches(0.8))
    tf = txBox.text_frame
    p = tf.paragraphs[0]
    p.text = "Relatório Executivo Mensal — Loja Hawara"
    p.font.size = Pt(24)
    p.font.bold = True
    
    # Inserir a imagem do gráfico gerado pelo Python no slide
    left = Inches(0.8)
    top = Inches(1.5)
    width = Inches(7.5)
    slide.shapes.add_picture(caminho_grafico, left, top, width=width)
    
    # Salvar o arquivo final do PowerPoint
    caminho_saida = "relatorio_gerencial_hawara.pptx"
    prs.save(caminho_saida)
    print(f"Apresentação executiva gerada com sucesso em: {caminho_saida}")

if __name__ == "__main__":
    gerar_relatorio_com_grafico()
