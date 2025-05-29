from PyPDF2 import PdfReader, PdfWriter
from fpdf import FPDF
from PIL import Image
import os


base_pasta_imagens = "04_correcao_completa"
base_pasta_pdfs = "05_PDF_resultados_do_aluno"
base_pasta_saida = "06_PDF_aluno_unico"
turmas = ["1 Ano A", "1 Ano B", "2 Ano A", "2 Ano B", "3 Ano A", "3 Ano B"]

os.makedirs(base_pasta_saida, exist_ok=True)

def redimensionar_imagem(caminho_imagem, largura=800):
    imagem = Image.open(caminho_imagem)
    fator_reducao = largura / imagem.width
    nova_altura = int(imagem.height * fator_reducao)
    imagem_redimensionada = imagem.resize((largura, nova_altura), Image.LANCZOS)  # Usando LANCZOS no lugar de ANTIALIAS
    caminho_temporario = caminho_imagem.replace(".png", "_compress.png")
    imagem_redimensionada.save(caminho_temporario, format="PNG", optimize=True, quality=50)
    return caminho_temporario

def combinar_imagem_com_pdf(nome_aluno, turma):
    pasta_imagens = os.path.join(base_pasta_imagens, turma)
    pasta_pdfs = os.path.join(base_pasta_pdfs, turma)
    pasta_saida = os.path.join(base_pasta_saida, turma)
    os.makedirs(pasta_saida, exist_ok=True)

    caminho_imagem = os.path.join(pasta_imagens, f"{nome_aluno}.png")
    caminho_pdf_notas = os.path.join(pasta_pdfs, f"{nome_aluno}.pdf")
    caminho_pdf_saida = os.path.join(pasta_saida, f"{nome_aluno}.pdf")

    if not os.path.exists(caminho_imagem) or not os.path.exists(caminho_pdf_notas):
        print(f"Imagem ou PDF de notas não encontrado para o aluno: {nome_aluno} na turma {turma}")
        return

    caminho_imagem_reduzida = redimensionar_imagem(caminho_imagem)

    pdf_imagem = FPDF()
    pdf_imagem.add_page()
    pdf_imagem.image(caminho_imagem_reduzida, x=10, y=10, w=pdf_imagem.w - 20)
    caminho_pdf_imagem = os.path.join(pasta_saida, f"{nome_aluno}_temp.pdf")
    pdf_imagem.output(caminho_pdf_imagem)

    pdf_writer = PdfWriter()
    pdf_imagem_reader = PdfReader(caminho_pdf_imagem)
    pdf_notas_reader = PdfReader(caminho_pdf_notas)

    pdf_writer.add_page(pdf_imagem_reader.pages[0])

    for page in pdf_notas_reader.pages:
        pdf_writer.add_page(page)

    with open(caminho_pdf_saida, "wb") as arquivo_saida:
        pdf_writer.write(arquivo_saida)

    os.remove(caminho_pdf_imagem)
    os.remove(caminho_imagem_reduzida)
    print(f"PDF combinado salvo para o aluno: {nome_aluno} na turma {turma}")

def gerar_pdf_unico_turma(turma):
    pasta_saida_turma = os.path.join(base_pasta_saida, turma)
    pdf_writer_turma = PdfWriter()
    caminho_pdf_turma = os.path.join(base_pasta_saida, f"{turma}_completo.pdf")

    arquivos_pdf_alunos = [f for f in os.listdir(pasta_saida_turma) if f.endswith(".pdf")]

    for arquivo_pdf in sorted(arquivos_pdf_alunos):
        caminho_pdf_aluno = os.path.join(pasta_saida_turma, arquivo_pdf)
        pdf_aluno_reader = PdfReader(caminho_pdf_aluno)

        for page in pdf_aluno_reader.pages:
            pdf_writer_turma.add_page(page)

    with open(caminho_pdf_turma, "wb") as arquivo_saida_turma:
        pdf_writer_turma.write(arquivo_saida_turma)
    print(f"PDF único gerado para a turma: {turma}")

for turma in turmas:
    pasta_imagens = os.path.join(base_pasta_imagens, turma)
    pasta_pdfs = os.path.join(base_pasta_pdfs, turma)

    imagens = [f.replace(".png", "") for f in os.listdir(pasta_imagens) if f.endswith(".png")]
    pdfs = [f.replace(".pdf", "") for f in os.listdir(pasta_pdfs) if f.endswith(".pdf")]

    for nome_aluno in imagens:
        if nome_aluno in pdfs:
            combinar_imagem_com_pdf(nome_aluno, turma)
        else:
            print(f"Imagem e PDF não correspondem para o aluno: {nome_aluno} na turma {turma}")

    gerar_pdf_unico_turma(turma)
