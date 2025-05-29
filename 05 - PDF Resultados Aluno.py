from PIL import Image, ImageDraw, ImageFilter, ImageFont
import plotly.graph_objects as go
from fpdf import FPDF
import os
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patheffects as patheffects

materias_por_turma = {
    "1 Ano A": [
        ("FÍSICA I", "1 a 4"),
        ("FÍSICA II", "5 a 8"),
        ("QUÍMICA I", "9 a 12"),
        ("QUÍMICA II", "13 a 16"),
        ("BIOLOGIA I", "17 a 20"),
        ("BIOLOGIA II", "21 a 24"),
        ("GEOGRAFIA", "25 a 28"),
        ("HISTÓRIA GERAL", "29 a 32"),
        ("HISTÓRIA DO BRASIL", "33 a 35"),
        ("FILOSOFIA", "36 a 38"),
        ("SOCIOLOGIA", "39 a 41"),
        ("L. PORTUGUESA", "42 a 46"),
        ("LITERATURA", "47 a 49"),
        ("I. CIENTÍFICA", "50 a 52"),
        ("L. INGLESA", "53 a 56"),
        ("ED. FÍSICA", "57 a 58"),
        ("MATEMÁTICA I", "59 a 63"),
        ("MATEMÁTICA II", "64 a 68"),
        ("ARTE", "69 a 70")
    ],
    "1 Ano B": [
        ("FÍSICA I", "1 a 4"),
        ("FÍSICA II", "5 a 8"),
        ("QUÍMICA I", "9 a 12"),
        ("QUÍMICA II", "13 a 16"),
        ("BIOLOGIA I", "17 a 20"),
        ("BIOLOGIA II", "21 a 24"),
        ("GEOGRAFIA", "25 a 28"),
        ("HISTÓRIA GERAL", "29 a 32"),
        ("HISTÓRIA DO BRASIL", "33 a 35"),
        ("FILOSOFIA", "36 a 38"),
        ("SOCIOLOGIA", "39 a 41"),
        ("L. PORTUGUESA", "42 a 46"),
        ("LITERATURA", "47 a 49"),
        ("I. CIENTÍFICA", "50 a 52"),
        ("L. INGLESA", "53 a 56"),
        ("ED. FÍSICA", "57 a 58"),
        ("MATEMÁTICA I", "59 a 63"),
        ("MATEMÁTICA II", "64 a 68"),
        ("ARTE", "69 a 70")
    ],
    "2 Ano A": [
        ("FÍSICA I", "1 a 5"),
        ("FÍSICA II", "6 a 10"),
        ("QUÍMICA I", "11 a 15"),
        ("QUÍMICA II", "16 a 20"),
        ("BIOLOGIA I", "21 a 25"),
        ("BIOLOGIA II", "26 a 30"),
        ("GEOGRAFIA", "31 a 34"),
        ("HISTÓRIA GERAL", "35 a 38"),
        ("HISTÓRIA DO BRASIL", "39 a 41"),
        ("FILOSOFIA", "42 a 44"),
        ("SOCIOLOGIA", "45 a 47"),
        ("L. PORTUGUESA", "48 a 52"),
        ("LITERATURA", "53 a 55"),
        ("I. CIENTÍFICA", "56 a 58"),
        ("L. INGLESA", "59 a 62"),
        ("INT. TEXTO", "63 a 67"),
        ("MATEMÁTICA I", "66 a 70"),
        ("MATEMÁTICA II", "71 a 75")
    ],
    "2 Ano B": [
        ("FÍSICA I", "1 a 5"),
        ("FÍSICA II", "6 a 10"),
        ("QUÍMICA I", "11 a 15"),
        ("QUÍMICA II", "16 a 20"),
        ("BIOLOGIA I", "21 a 25"),
        ("BIOLOGIA II", "26 a 30"),
        ("GEOGRAFIA", "31 a 34"),
        ("HISTÓRIA GERAL", "35 a 38"),
        ("HISTÓRIA DO BRASIL", "39 a 41"),
        ("FILOSOFIA", "42 a 44"),
        ("SOCIOLOGIA", "45 a 47"),
        ("L. PORTUGUESA", "48 a 52"),
        ("LITERATURA", "53 a 55"),
        ("I. CIENTÍFICA", "56 a 58"),
        ("L. INGLESA", "59 a 62"),
        ("INT. TEXTO", "63 a 67"),
        ("MATEMÁTICA I", "66 a 70"),
        ("MATEMÁTICA II", "71 a 75")
    ],
    "3 Ano A": [
        ("FÍSICA I", "1 a 5"),
        ("FÍSICA II", "6 a 10"),
        ("QUÍMICA I", "11 a 15"),
        ("QUÍMICA II", "16 a 20"),
        ("BIOLOGIA I", "21 a 25"),
        ("BIOLOGIA II", "26 a 30"),
        ("GEOGRAFIA", "31 a 34"),
        ("HISTÓRIA GERAL", "35 a 38"),
        ("HISTÓRIA DO BRASIL", "39 a 42"),
        ("FILOSOFIA", "43 a 46"),
        ("SOCIOLOGIA", "47 a 50"),
        ("L. PORTUGUESA", "51 a 56"),
        ("LITERATURA", "57 a 60"),
        ("L. INGLESA", "61 a 64"),
        ("INT. TEXTO", "65 a 68"),
        ("MATEMÁTICA I", "69 a 74"),
        ("MATEMÁTICA II", "75 a 80")
    ],
    "3 Ano B": [
        ("FÍSICA I", "1 a 5"),
        ("FÍSICA II", "6 a 10"),
        ("QUÍMICA I", "11 a 15"),
        ("QUÍMICA II", "16 a 20"),
        ("BIOLOGIA I", "21 a 25"),
        ("BIOLOGIA II", "26 a 30"),
        ("GEOGRAFIA", "31 a 34"),
        ("HISTÓRIA GERAL", "35 a 38"),
        ("HISTÓRIA DO BRASIL", "39 a 42"),
        ("FILOSOFIA", "43 a 46"),
        ("SOCIOLOGIA", "47 a 50"),
        ("L. PORTUGUESA", "51 a 56"),
        ("LITERATURA", "57 a 60"),
        ("L. INGLESA", "61 a 64"),
        ("INT. TEXTO", "65 a 68"),
        ("MATEMÁTICA I", "69 a 74"),
        ("MATEMÁTICA II", "75 a 80")
    ],
}

def carregar_acertos_excel(nome_aluno, turma):
    caminho_arquivo_acertos = f"04_correcao_completa/{turma}/acertos_por_disciplina_{turma}.xlsx"
    df_acertos = pd.read_excel(caminho_arquivo_acertos)
    if 'Aluno' not in df_acertos.columns:
        print(f"A coluna 'Aluno' não foi encontrada no arquivo de acertos {caminho_arquivo_acertos}")
        return None
    linha_aluno_acertos = df_acertos[df_acertos['Aluno'] == nome_aluno]
    if linha_aluno_acertos.empty:
        print(f"Aluno {nome_aluno} não encontrado no arquivo de acertos {caminho_arquivo_acertos}")
        return None
    materias = materias_por_turma[turma]
    acertos_por_materia = {}
    for materia, _ in materias:
        if materia in df_acertos.columns:
            acertos_por_materia[materia] = linha_aluno_acertos[materia].values[0]
        else:
            acertos_por_materia[materia] = 0
    return acertos_por_materia

def carregar_notas_excel(nome_aluno, turma):
    caminho_arquivo = f"04_correcao_completa/{turma}/notas_por_disciplina_{turma}.xlsx"
    df = pd.read_excel(caminho_arquivo)
    if 'Aluno' not in df.columns:
        print(f"A coluna 'Aluno' não foi encontrada no arquivo {caminho_arquivo}")
        return None
    linha_aluno = df[df['Aluno'] == nome_aluno]
    if linha_aluno.empty:
        print(f"Aluno {nome_aluno} não encontrado no arquivo {caminho_arquivo}")
        return None
    materias = materias_por_turma[turma]
    notas_por_materia = {}
    for materia, _ in materias:
        coluna_nota = f"Nota {materia}"
        if coluna_nota in df.columns:
            notas_por_materia[materia] = linha_aluno[coluna_nota].values[0]
        else:
            notas_por_materia[materia] = 0
    return notas_por_materia

def carregar_totais_acertos_erros_percentual(nome_aluno, turma):
    caminho_arquivo_totais = f"04_correcao_completa/{turma}/totais_acertos_erros_{turma}.xlsx"
    df_totais = pd.read_excel(caminho_arquivo_totais)
    if 'Aluno' not in df_totais.columns or 'Total Acertos' not in df_totais.columns or \
       'Total Erros' not in df_totais.columns or 'Percentual de Acertos' not in df_totais.columns or \
       'Percentual de Erros' not in df_totais.columns:
        print(f"Algumas colunas necessárias não foram encontradas no arquivo {caminho_arquivo_totais}")
        return None, None, None, None
    linha_aluno = df_totais[df_totais['Aluno'] == nome_aluno]
    if linha_aluno.empty:
        print(f"Aluno {nome_aluno} não encontrado no arquivo {caminho_arquivo_totais}")
        return None, None, None, None
    acertos = linha_aluno['Total Acertos'].values[0]
    erros = linha_aluno['Total Erros'].values[0]
    percentual_acertos = linha_aluno['Percentual de Acertos'].values[0]
    percentual_erros = linha_aluno['Percentual de Erros'].values[0]
    return acertos, erros, percentual_acertos, percentual_erros

def criar_fundo_arredondado_com_sombra(largura, altura, cor_fundo, raio=30, sombra=True):
    img_sombra = Image.new('RGBA', (largura + 40, altura + 40), (0, 0, 0, 0))
    sombra_offset = (20, 20)
    if sombra:
        draw_sombra = ImageDraw.Draw(img_sombra)
        draw_sombra.rounded_rectangle(
            ((sombra_offset[0], sombra_offset[1]), (largura + sombra_offset[0], altura + sombra_offset[1])),
            radius=raio,
            fill=(0, 0, 0, 180)
        )
        img_sombra = img_sombra.filter(ImageFilter.GaussianBlur(10))
    img_fundo = Image.new('RGBA', (largura, altura), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img_fundo)
    draw.rounded_rectangle(((0, 0), (largura, altura)), radius=raio, fill=cor_fundo)
    img_sombra.paste(img_fundo, sombra_offset, mask=img_fundo)
    return img_sombra

def obter_materias_turma():
    turma = input("Qual turma você quer gerar o PDF? (Escolha entre: 1 Ano A, 1 Ano B, 2 Ano A, 2 Ano B, 3 Ano A, 3 Ano B): ")
    if turma in materias_por_turma:
        return materias_por_turma[turma]
    else:
        print("Turma inválida! Por favor, tente novamente.")
        return obter_materias_turma()

def gerar_grafico_com_fundo(materia, intervalo, nota, max_valor=10):
    fig = go.Figure(go.Indicator(
        mode="gauge+number",
        value=nota,
        number={'valueformat': ".1f"},
        gauge={
            'axis': {'range': [0, max_valor]},
            'bar': {'color': "#fdaf00"},
            'steps': [
                {'range': [0, max_valor * 0.5], 'color': "#004b8d"},
                {'range': [max_valor * 0.5, max_valor], 'color': "#004b8d"}
            ],
            'threshold': {
                'line': {'color': "white", 'width': 4},
                'thickness': 0.75,
                'value': nota
            }
        }
    ))
    fig.update_layout(
        paper_bgcolor='rgba(255,255,255,0)',
        plot_bgcolor='rgba(255,255,255,0)',
        margin=dict(l=30, r=30, t=30, b=30),
        width=250,
        height=180
    )
    fig.write_image("grafico_temp.png")
    fundo_img = criar_fundo_arredondado_com_sombra(350, 270, cor_fundo="white", raio=50, sombra=True)
    grafico_img = Image.open("grafico_temp.png")
    posicao_x = (fundo_img.width - grafico_img.width) // 2
    posicao_y = 90
    fundo_img.paste(grafico_img, (posicao_x, posicao_y), grafico_img)
    draw = ImageDraw.Draw(fundo_img)
    font_titulo = ImageFont.truetype("arial.ttf", 30)
    font_subtitulo = ImageFont.truetype("arial.ttf", 20)
    texto_superior_bbox = draw.textbbox((0, 0), materia, font=font_titulo)
    texto_superior_width = texto_superior_bbox[2] - texto_superior_bbox[0]
    draw.text(((fundo_img.width - texto_superior_width) // 2, 40), materia, font=font_titulo, fill="black")
    draw.line((30, 85, fundo_img.width - 30, 85), fill="gray", width=2)
    texto_inferior_bbox = draw.textbbox((0, 0), intervalo, font=font_subtitulo)
    texto_inferior_width = texto_inferior_bbox[2] - texto_inferior_bbox[0]
    draw.text(((fundo_img.width - texto_inferior_width) // 2, fundo_img.height - 50), intervalo, font=font_subtitulo, fill="black")
    return fundo_img

def obter_turma_por_numero():
    opcoes_turma = {
        1: "1 Ano A",
        2: "1 Ano B",
        3: "2 Ano A",
        4: "2 Ano B",
        5: "3 Ano A",
        6: "3 Ano B"
    }
    print("Escolha a turma:")
    for numero, turma in opcoes_turma.items():
        print(f"{numero}: {turma}")
    while True:
        try:
            escolha = int(input("Digite o número da turma: "))
            if escolha in opcoes_turma:
                return opcoes_turma[escolha]
            else:
                print("Opção inválida. Tente novamente.")
        except ValueError:
            print("Entrada inválida. Por favor, digite um número.")

def gerar_grafico_desempenho(turma, nome_aluno, acertos_por_materia):
    dados = materias_por_turma[turma]
    materias = [materia for materia, intervalo in dados]
    maximos = []
    for materia, intervalo in dados:
        inicio, fim = intervalo.split(" a ")
        max_val = int(fim) - int(inicio) + 1
        maximos.append(max_val)
    acertos = [acertos_por_materia.get(materia, 0) for materia in materias]
    plt.figure(figsize=(12, 14))
    plt.barh(materias, maximos, color="#f0f0f0", edgecolor='black')
    plt.barh(materias, acertos, color="#004b8d")
    for i, (atual, max_val) in enumerate(zip(acertos, maximos)):
        plt.text(atual + 0.1, i, f'{atual} / {max_val}', va='center', color='black', fontsize=14)
    plt.yticks(fontsize=17)
    plt.title('Desempenho por Matéria', fontsize=25)
    plt.xlabel('Quantidade de Acertos', fontsize=25)
    plt.ylabel('Matéria', fontsize=25)
    plt.tight_layout()
    grafico_path = 'grafico_desempenho.png'
    plt.savefig(grafico_path)
    plt.close()
    return grafico_path

def gerar_grafico_donut_matplotlib(percentual_acertos, percentual_erros, dist_labels=1.2, title_fontsize=25, perc_fontsize=14, perc_color='black', figure_size=(6, 6)):
    labels = ['Acertos', 'Erros']
    valores = [percentual_acertos, percentual_erros]
    cores = ['#004b8d', '#fdaf00']
    explode = (0.1, 0)
    fig, ax = plt.subplots(figsize=figure_size)
    wedges, texts, autotexts = ax.pie(valores, colors=cores, labels=labels, autopct='%1.1f%%', explode=explode, startangle=90, pctdistance=0.85, labeldistance=dist_labels)
    centro_circulo = plt.Circle((0, 0), 0.70, fc='white')
    ax.add_artist(centro_circulo)
    plt.title('Percentual de Acertos e Erros', fontsize=title_fontsize)
    for autotext in autotexts:
        autotext.set_fontsize(perc_fontsize)
        autotext.set_color(perc_color)
        autotext.set_weight('bold')
    for text in texts:
        text.set_fontsize(16)
        text.set_color('black')
    plt.tight_layout()
    grafico_donut_path = 'grafico_donut_estilo_matplotlib.png'
    plt.savefig(grafico_donut_path)
    plt.close()
    return grafico_donut_path

def grafico_barras_acertos_erros(acertos, erros):
    categorias = ['Acertos', 'Erros']
    valores = [acertos, erros]
    fig, ax = plt.subplots(figsize=(6, 4))
    cores = ['#004b8d', '#fdaf00']
    barras = ax.bar(categorias, valores, color=cores, edgecolor='black', linewidth=1.5, zorder=3)
    for barra in barras:
        barra.set_path_effects([patheffects.withSimplePatchShadow()])
    ax.set_title('ERROS X ACERTOS', fontsize=20, pad=20)
    ax.set_ylabel('Quantidade', fontsize=12)
    for i, valor in enumerate(valores):
        ax.text(i, valor + 0.5, str(valor), ha='center', va='bottom', fontsize=12, fontweight='bold')
    ax.spines['top'].set_visible(False)
    ax.spines['right'].set_visible(False)
    ax.spines['left'].set_color('gray')
    ax.spines['bottom'].set_color('gray')
    ax.grid(True, axis='y', color='gray', linestyle='--', linewidth=0.5, alpha=0.7, zorder=0)
    grafico_barras_path = 'grafico_barras_acertos_erros.png'
    plt.tight_layout()
    plt.savefig(grafico_barras_path)
    plt.close()
    return grafico_barras_path

def gerar_pdf_com_graficos(altura_inicial, nome_aluno, turma, notas_por_materia, acertos_por_materia, acertos, erros, percentual_acertos, percentual_erros):
    materias = materias_por_turma[turma]
    largura_A4, altura_A4 = (595, 842)
    margem = 28.35
    espacamento_x = 5
    espacamento_y = 5
    largura_grafico = 110
    altura_grafico = 90
    espacamento_lateral = (largura_A4 - (5 * largura_grafico) - (4 * espacamento_x)) / 2
    pdf = FPDF(orientation='P', unit='pt', format=[largura_A4, altura_A4])
    pdf.add_page()
    pdf.image('logo.png', x=margem, y=margem, w=180)
    pdf.set_xy(margem + 190, margem + 10)
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(0, 10, "1º SIMULADO SESC ENSINO MÉDIO", ln=True, align='L')
    pdf.set_xy(margem + 190, margem + 30)
    pdf.set_font('Arial', '', 10)
    pdf.cell(0, 10, "RESULTADOS", ln=True, align='L')
    pdf.set_xy(margem, margem + 60)
    pdf.set_font('Arial', '', 15)
    nome_aluno_sem_extensao = nome_aluno.replace(".png", "")
    pdf.cell(0, 10, nome_aluno_sem_extensao, ln=True, align='C')
    for index, (materia, intervalo) in enumerate(materias):
        i = index // 5
        j = index % 5
        pos_x = espacamento_lateral + j * (largura_grafico + espacamento_x)
        pos_y = altura_inicial + i * (altura_grafico + espacamento_y)
        nota = notas_por_materia.get(materia, 0)
        img = gerar_grafico_com_fundo(
            materia=materia,
            intervalo=intervalo,
            nota=nota,
            max_valor=10
        )
        img_path = f'grafico_{index + 1}.png'
        img.save(img_path)
        pdf.image(img_path, pos_x, pos_y, largura_grafico, altura_grafico)
        os.remove(img_path)
    grafico_desempenho_path = gerar_grafico_desempenho(turma, nome_aluno, acertos_por_materia)
    pdf.image(grafico_desempenho_path, x=margem, y=altura_A4 - 330, w=250)
    grafico_donut_path = gerar_grafico_donut_matplotlib(percentual_acertos, percentual_erros)
    pdf.image(grafico_donut_path, x=margem + 320, y=altura_A4 - 330, w=150, h=150)
    grafico_barras_path = grafico_barras_acertos_erros(acertos, erros)
    pdf.image(grafico_barras_path, x=margem + 280, y=altura_A4 - 150, w=220, h=120)
    pasta_pdf = f"05_PDF_resultados_do_aluno/{turma}"
    if not os.path.exists(pasta_pdf):
        os.makedirs(pasta_pdf)
    nome_aluno_sem_extensao = os.path.splitext(nome_aluno)[0]
    pdf.output(f"{pasta_pdf}/{nome_aluno_sem_extensao}.pdf")

def main():
    continuar = True
    while continuar:
        turma = obter_turma_por_numero()
        caminho_arquivo = f"04_correcao_completa/{turma}/notas_por_disciplina_{turma}.xlsx"
        df = pd.read_excel(caminho_arquivo)
        for nome_aluno in df['Aluno']:
            notas_por_materia = carregar_notas_excel(nome_aluno, turma)
            acertos_por_materia = carregar_acertos_excel(nome_aluno, turma)
            acertos, erros, percentual_acertos, percentual_erros = carregar_totais_acertos_erros_percentual(nome_aluno, turma)
            if notas_por_materia and acertos_por_materia and acertos is not None:
                gerar_pdf_com_graficos(
                    altura_inicial=120,
                    nome_aluno=nome_aluno,
                    turma=turma,
                    notas_por_materia=notas_por_materia,
                    acertos_por_materia=acertos_por_materia,
                    acertos=acertos,
                    erros=erros,
                    percentual_acertos=percentual_acertos,
                    percentual_erros=percentual_erros
                )
        resposta = input("Deseja corrigir outra turma? (s/n): ").strip().lower()
        if resposta != 's':
            continuar = False

if __name__ == "__main__":
    main()
