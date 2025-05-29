import os
import pandas as pd
from fpdf import FPDF
import qrcode
import json

class GabaritoPDF(FPDF):
    def __init__(self, total_questoes, titulo, subtitulo, nome_aluno, turma_aluno, codigo_aluno, sala_aluno, carteira_aluno, data_simulado, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.total_questoes = total_questoes
        self.titulo = titulo
        self.subtitulo = subtitulo
        self.nome_aluno = nome_aluno
        self.turma_aluno = turma_aluno
        self.codigo_aluno = codigo_aluno
        self.sala_aluno = sala_aluno
        self.carteira_aluno = carteira_aluno
        self.data_simulado = data_simulado

    def header(self):
        self.set_fill_color(0, 0, 0)
        square_size = 5
        margin = 10
        self.rect(margin, margin, square_size, square_size, 'F')
        self.rect(210 - margin - square_size, margin, square_size, square_size, 'F')
        self.rect(margin, 297 - margin - square_size, square_size, square_size, 'F')
        self.rect(210 - margin - square_size, 297 - margin - square_size, square_size, square_size, 'F')
        self.image('logo.png', 15, 15, 70)
        self.set_xy(90, 20)
        self.set_font('Arial', 'B', 12)
        self.cell(0, 5, self.titulo, ln=True, align='L')
        self.set_xy(90, 25)
        self.set_font('Arial', '', 10)
        self.cell(0, 5, self.subtitulo, ln=True, align='L')
        self.rect(12, 40, 90, 30)
        self.set_xy(14, 42)
        self.set_font('Arial', '', 10)
        self.cell(0, 8, f'NOME: {self.nome_aluno}', ln=True, align='L')
        self.set_xy(14, 50)
        self.cell(0, 8, f'TURMA: {self.turma_aluno}', ln=True, align='L')
        self.set_xy(14, 58)
        self.cell(0, 8, f'CODIGO: {self.codigo_aluno}', ln=True, align='L')
        self.rect(104, 40, 50, 30)
        self.set_xy(104, 42)
        self.set_font('Arial', 'B', 10)
        self.cell(50, 8, 'ENSALAMENTO', ln=True, align='C')
        self.set_xy(104, 50)
        self.set_font('Arial', '', 10)
        self.cell(50, 8, f' {self.sala_aluno}', ln=True, align='C')
        self.set_xy(104, 58)
        self.cell(50, 8, f'CARTEIRA {self.carteira_aluno}', ln=True, align='C')
        self.line(60, 103, 150, 103)
        self.set_xy(0, 104)
        self.set_font('Arial', '', 10)
        self.cell(0, 10, 'ASSINATURA DO PARTICIPANTE:', ln=True, align='C')
        self.set_xy(15, 75)
        self.set_font('Arial', 'B', 10)
        self.cell(0, 6, 'INSTRUÇÕES', ln=True, align='C')
        self.set_font('Arial', '', 8)
        self.cell(0, 6, 'Preencha completamente a opção escolhida sem rasuras.', ln=True, align='C')
        self.rect(12, 72, 180, 20)
        self.adicionar_qr_code(self.nome_aluno, self.codigo_aluno, self.turma_aluno, self.data_simulado, self.total_questoes, self.titulo)

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

    def desenha_gabarito(self):
        self.set_font('Arial', '', 7)
        self.set_fill_color(0, 0, 0)
        self.set_line_width(0.2)

        start_x = 10
        start_y = 120
        circle_diameter = 3.5
        gap_between_circles = 6
        gap_between_lines = 6
        questions_per_column = 25
        column_width = 45
        alternativas = ['A', 'B', 'C', 'D', 'E']

        for question in range(1, 101):
            column = (question - 1) // questions_per_column
            y = start_y + ((question - 1) % questions_per_column) * gap_between_lines
            x = start_x + column * column_width

            if question <= self.total_questoes:
                self.set_xy(x, y)
                self.cell(15, 10, f'{question} -', 0, 0, 'R')
                for i in range(5):
                    cx = x + 15 + i * gap_between_circles
                    cy = y + 2.5
                    self.ellipse(cx, cy, circle_diameter, circle_diameter)
                    self.set_text_color(169, 169, 169)
                    self.set_xy(cx, cy)
                    self.cell(circle_diameter, circle_diameter, alternativas[i], 0, 0, 'C')
            else:
                self.set_text_color(200, 200, 200)
                self.set_xy(x, y)
                self.cell(15, 10, "", 0, 0, 'R')
                for i in range(5):
                    cx = x + 15 + i * gap_between_circles
                    cy = y + 2.5
                    self.set_xy(cx, cy)
                    self.cell(circle_diameter, circle_diameter, "-", 0, 0, 'C')
            self.set_text_color(0, 0, 0)

def ler_dados_excel(caminho_excel):
    df = pd.read_excel(caminho_excel)
    return df

def definir_total_questoes(turma):
    turma = turma.strip()
    if turma in ['1 Ano A', '1 Ano B']:
        return 70
    elif turma in ['2 Ano A', '2 Ano B']:
        return 75
    elif turma in ['3 Ano A', '3 Ano B']:
        return 80
    else:
        raise ValueError(f"Turma desconhecida: {turma}")

def gerar_pdf_unico(df, titulo, subtitulo, nome_arquivo_final, data_simulado):
    pdf = GabaritoPDF(0, titulo, subtitulo, '', '', '', '', '', data_simulado, 'P', 'mm', 'A4')
    for index, row in df.iterrows():
        nome_aluno = row['Nome']
        turma_aluno = row['Turma']
        codigo_aluno = row['Código']
        sala_aluno = row['Sala']
        carteira_aluno = row['Carteira']

        if 'Questões' in row and not pd.isnull(row['Questões']):
            total_questoes = int(row['Questões'])
        else:
            total_questoes = definir_total_questoes(turma_aluno.strip())

        pdf.total_questoes = total_questoes
        pdf.nome_aluno = nome_aluno
        pdf.turma_aluno = turma_aluno
        pdf.codigo_aluno = codigo_aluno
        pdf.sala_aluno = sala_aluno
        pdf.carteira_aluno = carteira_aluno
        pdf.add_page()
        pdf.desenha_gabarito()

    pdf.output(nome_arquivo_final)

def escolher_opcao():
    print("Selecione a opção:")
    opcoes = ['1 Ano A', '1 Ano B', '2 Ano A', '2 Ano B', '3 Ano A', '3 Ano B', '2ª Chamada', 'Adaptado']
    for i, opcao in enumerate(opcoes, 1):
        print(f"{i}. {opcao}")
    escolha = int(input("Digite o número da opção: "))
    return opcoes[escolha - 1]

def perguntar_gerar_novamente():
    while True:
        resposta = input("Você deseja gerar mais algum gabarito? (S/N): ").strip().upper()
        if resposta == 'S':
            return True
        elif resposta == 'N':
            return False
        else:
            print("Resposta inválida. Digite 'S' para Sim ou 'N' para Não.")


titulo = "4º SIMULADO SESC ENSINO MÉDIO"
subtitulo = "1ª CHAMADA | 17/05/2025"
data_simulado = "17/05/2025"

continuar = True
while continuar:
    opcao_escolhida = escolher_opcao()

    if opcao_escolhida == '2ª Chamada':
        caminho_excel = '01_gerar_pdf/2a_chamada.xlsx'
        nome_arquivo_final = '01_gerar_pdf/Gabaritos_2a_Chamada.pdf'
        df_alunos = ler_dados_excel(caminho_excel)
        gerar_pdf_unico(df_alunos, titulo, subtitulo, nome_arquivo_final, data_simulado)

    elif opcao_escolhida == 'Adaptado':
        print("Opção Adaptado selecionada. Por favor, digite os dados manualmente.")
        nome_aluno = input("Digite o nome do aluno: ")
        codigo_aluno = input("Digite o código do aluno: ")
        carteira_aluno = input("Digite a carteira do aluno: ")
        sala_aluno = input("Digite a sala do aluno: ")
        turma_aluno = input("Digite a turma do aluno: ")  # Agora solicita a turma manualmente
        total_questoes = int(input("Digite a quantidade de questões do gabarito: "))

        data = {
            'Nome': [nome_aluno],
            'Turma': [turma_aluno],
            'Código': [codigo_aluno],
            'Sala': [sala_aluno],
            'Carteira': [carteira_aluno],
            'Questões': [total_questoes]
        }
        df_alunos = pd.DataFrame(data)
        nome_arquivo_final = '01_gerar_pdf/Gabarito_Adaptado.pdf'
        gerar_pdf_unico(df_alunos, titulo, subtitulo, nome_arquivo_final, data_simulado)

    else:
        caminho_excel = f'01_gerar_pdf/{opcao_escolhida}.xlsx'
        nome_arquivo_final = f'01_gerar_pdf/Gabaritos_{opcao_escolhida}.pdf'
        df_alunos = ler_dados_excel(caminho_excel)
        gerar_pdf_unico(df_alunos, titulo, subtitulo, nome_arquivo_final, data_simulado)

    print(f"PDF gerado com sucesso: {nome_arquivo_final}")
    continuar = perguntar_gerar_novamente()

print("Processo concluído.")
