from fpdf import FPDF

class GabaritoPDF(FPDF):
    def __init__(self, total_questoes, titulo, subtitulo, gabarito_oficial=None, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.total_questoes = total_questoes
        self.titulo = titulo
        self.subtitulo = subtitulo
        self.gabarito_oficial = gabarito_oficial if gabarito_oficial else [''] * total_questoes

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
        self.cell(0, 8, 'NOME: Mascara do Simulado', ln=True, align='L')
        self.set_xy(14, 50)
        self.cell(0, 8, 'TURMA: ', ln=True, align='L')
        self.set_xy(14, 58)
        self.cell(0, 8, 'CODIGO: ', ln=True, align='L')
        self.rect(104, 40, 50, 30)
        self.set_xy(104, 42)
        self.set_font('Arial', 'B', 10)
        self.cell(50, 8, 'ENSALAMENTO', ln=True, align='C')
        self.set_xy(104, 50)
        self.set_font('Arial', '', 10)
        self.cell(50, 8, 'SALA', ln=True, align='C')
        self.set_xy(104, 58)
        self.cell(50, 8, 'CARTEIRA', ln=True, align='C')
        self.line(60, 103, 150, 103)
        self.set_xy(0, 104)
        self.set_font('Arial', '', 10)
        self.cell(0, 10, 'ASSINATURA DO PARTICIPANTE:', ln=True, align='C')
        self.set_xy(15, 75)
        self.set_font('Arial', 'B', 10)
        self.cell(0, 6, 'INSTRUÇÕES', ln=True, align='C')
        self.set_font('Arial', '', 8)
        self.cell(0, 6, 'Preencha totalmente a opção escolhida conforme a imagem ao lado.', ln=True, align='C')
        self.rect(12, 72, 180, 20)

    def desenha_gabarito(self):
        self.set_font('Arial', '', 7)
        self.set_fill_color(0, 0, 0)
        self.set_line_width(0.2)
        start_x = 10
        start_y = 120
        circle_diameter = 3.5
        gap_between_circles = 6
        gap_between_lines = 6
        num_columns = 6
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
                resposta_correta = self.gabarito_oficial[question - 1]
                for i in range(5):
                    cx = x + 15 + i * gap_between_circles
                    cy = y + 2.5
                    if resposta_correta == '0' or alternativas[i] == resposta_correta:
                        self.ellipse(cx, cy, circle_diameter, circle_diameter, 'F')
                    else:
                        self.ellipse(cx, cy, circle_diameter, circle_diameter)
                        self.set_text_color(169, 169, 169)
                        self.set_xy(cx, cy)
                        self.cell(circle_diameter, circle_diameter, alternativas[i], 0, 0, 'C')
            else:
                self.set_xy(x, y)
                self.cell(15, 10, "", 0, 0, 'R')
                for i in range(5):
                    cx = x + 15 + i * gap_between_circles
                    cy = y + 2.5
                    self.set_xy(cx, cy)
                    self.cell(circle_diameter, circle_diameter, "-", 0, 0, 'C')
            self.set_text_color(0, 0, 0)


def escolher_turma():
    print("Selecione a turma:")
    turmas = ['1 Ano A', '1 Ano B', '2 Ano A', '2 Ano B', '3 Ano A', '3 Ano B']
    for i, turma in enumerate(turmas, 1):
        print(f"{i}. {turma}")
    escolha = int(input("Digite o número da turma: "))
    return turmas[escolha - 1]


def definir_total_questoes(turma):
    if turma in ['1 Ano A', '1 Ano B']:
        return 70
    elif turma in ['2 Ano A', '2 Ano B']:
        return 75
    elif turma in ['3 Ano A', '3 Ano B']:
        return 80

def obter_gabarito(total_questoes):
    while True:
        gabarito = input(f"Digite o gabarito oficial ({total_questoes} questões, separadas por espaço): ").upper().split()
        if len(gabarito) == total_questoes and all(resposta in ['A', 'B', 'C', 'D', 'E', '0'] for resposta in gabarito):
            return gabarito
        else:
            print(f"Quantidade de respostas fornecidas: {len(gabarito)}. Deve conter {total_questoes} respostas válidas. Tente novamente.")


def perguntar_gerar_mais():
    resposta = input("Você deseja gerar mais um gabarito? (S/N): ").upper()
    return resposta == 'S'


continuar = True
while continuar:
    turma = escolher_turma()
    total_questoes = definir_total_questoes(turma)

    print(f"Digite o gabarito oficial para {turma} com {total_questoes} questões.")
    gabarito_oficial = obter_gabarito(total_questoes)
    titulo = f"Gabarito - {turma}"
    subtitulo = "Simulado Oficial"
    pdf = GabaritoPDF(total_questoes, titulo, subtitulo, gabarito_oficial)
    pdf.add_page()
    pdf.desenha_gabarito()
    nome_arquivo = f"02_gerar_mascara/Gabarito_{turma}.pdf"
    pdf.output(nome_arquivo)

    print(f"Gabarito salvo com sucesso: {nome_arquivo}")
    continuar = perguntar_gerar_mais()

print("Processo concluído.")
