import os
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from fpdf import FPDF
import qrcode
import json


largura_bolinha = 35
altura_bolinha = 35
colunas = 4
linhas_por_coluna = 25

posicoes_colunas_aluno = [132, 535, 937, 1339]
posicoes_colunas_gabarito = [132, 535, 937, 1339]
posicoes_colunas_secundarias = [0, 53, 107, 160, 213]

alturas_linhas_aluno = [
    977, 1029, 1081, 1133, 1185, 1237, 1287, 1339, 1391, 1443,
    1493, 1545, 1596, 1645, 1696, 1748, 1800, 1850, 1900, 1952,
    2003, 2055, 2106, 2158, 2210
]

alturas_linhas_gabarito = [
    964, 1016, 1068, 1120, 1172, 1224, 1274, 1326, 1378, 1430,
    1480, 1532, 1584, 1636, 1688, 1740, 1792, 1842, 1894, 1946,
    1997, 2049, 2100, 2152, 2204
]

class GabaritoPDF(FPDF):
    def __init__(self, total_questoes, simulado_number, professor_info, turma_info, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.total_questoes = total_questoes
        self.simulado_number = simulado_number
        self.professor_info = professor_info
        self.turma_info = turma_info

    def header(self):
        self.image('logo.png', 15, 15, 70)
        self.set_xy(90, 20)
        self.set_font('Arial', 'B', 12)
        self.cell(0, 5, f"{self.simulado_number}º SIMULADO SESC ENSINO MÉDIO", ln=True, align='L')
        self.set_xy(90, 25)
        self.set_font('Arial', '', 10)
        self.cell(0, 5, "RESULTADOS", ln=True, align='L')
        self.rect(12, 40, 180, 13)
        self.set_xy(12, 40)
        self.set_font('Arial', 'B', 10)
        self.cell(180, 6, self.professor_info, ln=True, align='C')
        self.set_font('Arial', '', 8)
        self.cell(180, 6, self.turma_info, ln=True, align='C')

    def adicionar_qr_code(self, nome_aluno, codigo_aluno, turma_aluno, data_simulado, total_questoes, titulo_simulado):
        dados_qr = {
            "nome": nome_aluno,
            "codigo": codigo_aluno,
            "turma": turma_aluno,
            "data": data_simulado,
            "questoes": total_questoes,
            "simulado": titulo_simulado
        }
        dados_qr_json = json.dumps(dados_qr)
        qr = qrcode.QRCode(version=1, box_size=10, border=4)
        qr.add_data(dados_qr_json)
        qr.make(fit=True)
        img_qr = qr.make_image(fill='black', back_color='white')
        qr_code_filename = f'qrcode_{nome_aluno}.png'
        img_qr.save(qr_code_filename)
        self.image(qr_code_filename, 158, 38, 33)
        if os.path.exists(qr_code_filename):
            os.remove(qr_code_filename)

def draw_table(pdf, title, headers, data, col_widths,
               header_bg_color=(0, 75, 141),
               row_bg_color=(255, 255, 255),
               alt_row_bg_color=(174, 210, 242)):
    row_height = 6  # Altura das linhas (em mm)
    if pdf.get_y() < 53:
        pdf.set_y(53)
    pdf.ln(5)
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(0, row_height+2, title, ln=True, align='C')
    pdf.ln(3)

    effective_width = pdf.w - pdf.l_margin - pdf.r_margin
    table_width = sum(col_widths)
    x_start = pdf.l_margin + (effective_width - table_width) / 2

    def print_headers():
        pdf.set_x(x_start)
        pdf.set_font('Arial', 'B', 10)
        pdf.set_fill_color(*header_bg_color)
        pdf.set_text_color(255, 255, 255)
        for i, header in enumerate(headers):
            pdf.cell(col_widths[i], row_height, header, border=1, align='C', fill=True)
        pdf.ln()
        pdf.set_text_color(0, 0, 0)
        pdf.set_font('Arial', '', 10)

    print_headers()

    for row in data:
        if pdf.get_y() + row_height > pdf.h - pdf.b_margin:
            pdf.add_page()
            pdf.set_y(58)
            print_headers()
        pdf.set_x(x_start)
        if data.index(row) % 2 == 0:
            pdf.set_fill_color(*row_bg_color)
        else:
            pdf.set_fill_color(*alt_row_bg_color)
        for i, cell in enumerate(row):
            pdf.cell(col_widths[i], row_height, str(cell), border=1, align='C', fill=True)
        pdf.ln()

def gerar_grafico_erros_acertos_por_questao(question_data, output_filename):
    questions = [item[0] for item in question_data]
    acertos = [int(item[1]) for item in question_data]
    erros = [int(item[2]) for item in question_data]
    x = np.arange(len(questions))
    width = 0.35

    plt.figure(figsize=(8, 5))
    plt.bar(x - width / 2, acertos, width, label='Acertos', color='#004b8d')
    plt.bar(x + width / 2, erros, width, label='Erros', color='#fdaf00')
    plt.xlabel('Questão')
    plt.ylabel('Quantidade')
    plt.title('Erros e Acertos por Questão')
    plt.xticks(x, questions, rotation=45)
    plt.legend()
    plt.tight_layout()
    plt.savefig(output_filename)
    plt.close()
    return output_filename

col_widths_alf = [100, 15]
col_widths_desc = [20, 100, 15, 15, 15]
col_widths_mini = [20, 20, 20]

def processar_turma(pdf, turma_nome, excel_path):
    try:
        df = pd.read_excel(excel_path)
    except FileNotFoundError:
        print(f"Arquivo não encontrado para {turma_nome} em {excel_path}. Ignorando turma.")
        return

    pdf.turma_info = turma_nome

    df_table1 = df[["Aluno", "Nota"]].sort_values("Aluno")
    data_alf = [[row[0], f'{row[1]:.1f}'] for row in df_table1.values.tolist()]
    pdf.add_page()
    draw_table(pdf, "NOTAS - ORDEM ALFABÉTICA", ["Nome do Aluno", "Nota"], data_alf, col_widths_alf)

    df_table2 = df[["Aluno", "Acertos", "Nota", "Erros"]]
    df_table2 = df_table2.sort_values(by=["Acertos", "Aluno"], ascending=[False, True])
    data_desc = []
    for row in df_table2.values.tolist():
        ranking = f"{len(data_desc) + 1}º"
        data_desc.append([ranking, row[0], row[1], row[3], f'{row[2]:.1f}'])
    pdf.add_page()
    draw_table(pdf, "RANKING", ["Ranking", "Aluno", "Acertos", "Erros", "Nota"], data_desc, col_widths_desc)

    try:
        df_desempenho = pd.read_excel(excel_path, sheet_name="Desempenho por Questão")
    except FileNotFoundError:
        print(f"Aba 'Desempenho por Questão' não encontrada para {turma_nome} em {excel_path}. Ignorando seção.")
        return
    mini_table = df_desempenho[["Questão", "Acertos", "Erros"]]
    mini_table_data = mini_table.values.tolist()
    question_data = [[row[0], row[1], row[2]] for row in mini_table_data]
    pdf.add_page()
    draw_table(pdf, "DESEMPENHO POR QUESTÃO", ["Questão", "Acertos", "Erros"], mini_table_data, col_widths_mini)
    grafico_filename = f"grafico_erros_acertos_{turma_nome.replace(' ', '_')}.png"
    gerar_grafico_erros_acertos_por_questao(question_data, grafico_filename)
    x_img = (pdf.w - 120) / 2
    pdf.image(grafico_filename, x=x_img, y=pdf.get_y() + 10, w=120)

    try:
        df_acima_abaixo = pd.read_excel(excel_path, sheet_name="Acima-Abaixo Média")
    except FileNotFoundError:
        print(f"Aba 'Acima-Abaixo Média' não encontrada para {turma_nome} em {excel_path}. Ignorando seção.")
        return
    acima = df_acima_abaixo["Acima da Média"].iloc[0]
    abaixo = df_acima_abaixo["Abaixo da Média"].iloc[0]
    data_acima_abaixo = [["Acima da Média", "Abaixo da Média"], [acima, abaixo]]
    col_widths_acima = [40, 40]
    pdf.add_page()
    draw_table(pdf, "ACIMA/ABAIXO MÉDIA", ["Acima da Média", "Abaixo da Média"], data_acima_abaixo, col_widths_acima)
    labels = ["Acima da Média", "Abaixo da Média"]
    sizes = [acima, abaixo]
    plt.figure(figsize=(6,6))
    wedges, texts, autotexts = plt.pie(sizes, labels=labels, autopct='%1.1f%%', startangle=90, wedgeprops=dict(width=0.4),
                                       colors=['#004b8d', '#fdaf00'])
    plt.title("Alunos Acima/Abaixo da Média")
    plt.tight_layout()
    donut_filename = f"donut_acima_abaixo_{turma_nome.replace(' ', '_')}.png"
    plt.savefig(donut_filename)
    plt.close()
    x_img = (pdf.w - 100) / 2
    pdf.image(donut_filename, x=x_img, y=pdf.get_y() + 10, w=100)

grades = ["1 Ano A", "1 Ano B", "2 Ano A", "2 Ano B", "3 Ano A"]
disciplines = [
    ("FÍSICA I", 1, 4), ("FÍSICA II", 5, 8), ("QUÍMICA I", 9, 12),
    ("QUÍMICA II", 13, 16), ("BIOLOGIA I", 17, 20), ("BIOLOGIA II", 21, 24),
    ("GEOGRAFIA", 25, 28), ("HISTÓRIA GERAL", 29, 32), ("HISTÓRIA DO BRASIL", 33, 35),
    ("FILOSOFIA", 36, 38), ("SOCIOLOGIA", 39, 41),
    ("L. PORTUGUESA", 42, 46), ("LITERATURA", 47, 49),
    ("I. CIENTÍFICA", 50, 52), ("L. INGLESA", 53, 56),
    ("ED. FÍSICA", 57, 58), ("MATEMÁTICA I", 59, 63),
    ("MATEMÁTICA II", 64, 68), ("ARTE", 69, 70), ("INT. TEXTO", 69, 70)
]

destino = "06_PDF_resultados_do_professor"
if not os.path.exists(destino):
    os.makedirs(destino)

simulado_number = input("Digite o número do simulado: ").strip()

for discipline, start_idx, end_idx in disciplines:
    pdf = GabaritoPDF(
        total_questoes=75,
        simulado_number=simulado_number,
        professor_info=discipline,
        turma_info=""
    )
    for grade in grades:
        excel_path = os.path.join("04_correcao_completa", grade, f"{discipline}.xlsx")
        processar_turma(pdf, grade.upper(), excel_path)
    pdf_filename = os.path.join(destino, f"{discipline} - {simulado_number}º Simulado - Resultados.pdf")
    pdf.output(pdf_filename)
    print(f"PDF gerado: {pdf_filename}")
