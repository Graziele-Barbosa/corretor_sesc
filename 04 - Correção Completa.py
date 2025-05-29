import numpy as np
import os
import pandas as pd
import cv2


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

def detectar_bolinhas(imagem_caminho, posicoes_colunas_principais, alturas_linhas, num_questoes):
    imagem_alinhada = cv2.imread(imagem_caminho)
    imagem_cinza = cv2.cvtColor(imagem_alinhada, cv2.COLOR_BGR2GRAY)
    _, thresh = cv2.threshold(imagem_cinza, 150, 255, cv2.THRESH_BINARY_INV)
    kernel = np.ones((3, 3), np.uint8)
    thresh = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, kernel)
    bolinhas_detectadas = []
    for col in range(colunas):
        for linha in range(min(linhas_por_coluna, num_questoes - col * linhas_por_coluna)):
            x_base = posicoes_colunas_principais[col]
            y_base = alturas_linhas[linha]
            alternativas_marcadas = []
            for i, offset in enumerate(posicoes_colunas_secundarias):
                x_alternativa = x_base + offset
                y_alternativa = y_base
                roi = thresh[y_alternativa:y_alternativa + altura_bolinha,
                             x_alternativa:x_alternativa + largura_bolinha]
                total_pixels = roi.size
                pixels_preenchidos = cv2.countNonZero(roi)
                if total_pixels > 0 and (pixels_preenchidos / total_pixels) > 0.4:
                    alternativas_marcadas.append(i)
            if len(alternativas_marcadas) == 1:
                bolinhas_detectadas.append(alternativas_marcadas[0])
            elif len(alternativas_marcadas) == 5:
                bolinhas_detectadas.append("An")
            else:
                bolinhas_detectadas.append(None)
    return bolinhas_detectadas

def process_files(input_folder, output_folder, gabarito_path, intervalos_nomeados, num_questoes):
    bolinhas_gabarito = detectar_bolinhas(gabarito_path, posicoes_colunas_gabarito, alturas_linhas_gabarito, num_questoes)
    max_acertos_disciplina = {disciplina: (fim - inicio + 1) for disciplina, inicio, fim in intervalos_nomeados}
    resultados_acertos = {}
    discipline_question_stats = {}
    for intervalo in intervalos_nomeados:
        nome_materia, inicio, fim = intervalo
        inicio_idx = inicio - 1
        discipline_question_stats[nome_materia] = {q + 1: {'acertos': 0, 'erros': 0} for q in range(inicio_idx, min(fim, num_questoes))}
    global_question_stats = {i: {'acertos': 0, 'erros': 0} for i in range(1, num_questoes + 1)}
    for filename in os.listdir(input_folder):
        if filename.lower().endswith(('.png', '.jpg', '.jpeg')):
            file_path = os.path.join(input_folder, filename)
            bolinhas_aluno = detectar_bolinhas(file_path, posicoes_colunas_aluno, alturas_linhas_aluno, num_questoes)
            imagem_alinhada = cv2.imread(file_path)
            acertos_por_intervalo = {}
            for q in range(num_questoes):
                if bolinhas_gabarito[q] == "An" or (bolinhas_aluno[q] is not None and bolinhas_aluno[q] == bolinhas_gabarito[q]):
                    global_question_stats[q + 1]['acertos'] += 1
                else:
                    global_question_stats[q + 1]['erros'] += 1
            for intervalo in intervalos_nomeados:
                nome_materia, inicio, fim = intervalo
                inicio_idx = inicio - 1
                acertos = 0
                for q in range(inicio_idx, min(fim, num_questoes)):
                    question_number = q + 1
                    if bolinhas_gabarito[q] == "An" or (bolinhas_aluno[q] is not None and bolinhas_aluno[q] == bolinhas_gabarito[q]):
                        acertos += 1
                        discipline_question_stats[nome_materia][question_number]['acertos'] += 1
                    else:
                        discipline_question_stats[nome_materia][question_number]['erros'] += 1
                acertos_por_intervalo[nome_materia] = acertos
            resultados_acertos[filename] = acertos_por_intervalo
            for col in range(colunas):
                for linha in range(min(linhas_por_coluna, num_questoes - col * linhas_por_coluna)):
                    x_base = posicoes_colunas_aluno[col]
                    y_base = alturas_linhas_aluno[linha]
                    aluno_resposta = bolinhas_aluno[linha + col * linhas_por_coluna]
                    gabarito_resposta = bolinhas_gabarito[linha + col * linhas_por_coluna]
                    if aluno_resposta is None:
                        x_inicial = x_base + posicoes_colunas_secundarias[0]
                        x_final = x_base + posicoes_colunas_secundarias[-1] + largura_bolinha
                        cv2.rectangle(imagem_alinhada, (x_inicial, y_base), (x_final, y_base + altura_bolinha), (0, 0, 255), 2)
                        cv2.putText(imagem_alinhada, "X", (x_final + 10, y_base + altura_bolinha - 5),
                                    cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 0, 255), 2)
                    elif aluno_resposta == "An":
                        x_inicial = x_base + posicoes_colunas_secundarias[0]
                        x_final = x_base + posicoes_colunas_secundarias[-1] + largura_bolinha
                        cv2.rectangle(imagem_alinhada, (x_inicial, y_base), (x_final, y_base + altura_bolinha), (0, 255, 255), 2)
                        cv2.putText(imagem_alinhada, "An", (x_final + 10, y_base + altura_bolinha - 5),
                                    cv2.FONT_HERSHEY_SIMPLEX, 1, (0, 255, 255), 2)
                    else:
                        if isinstance(gabarito_resposta, int):
                            if aluno_resposta == gabarito_resposta:
                                x_alternativa_E = x_base + posicoes_colunas_secundarias[4]
                                cv2.putText(imagem_alinhada, "V",
                                            (x_alternativa_E + largura_bolinha + 5, y_base + altura_bolinha - 5),
                                            cv2.FONT_HERSHEY_SIMPLEX, 1, (34, 139, 34), 2)
                            else:
                                x_alternativa = x_base + posicoes_colunas_secundarias[gabarito_resposta]
                                cv2.rectangle(imagem_alinhada, (x_alternativa, y_base),
                                              (x_alternativa + largura_bolinha, y_base + altura_bolinha), (255, 0, 0), 2)
                                x_alternativa_E = x_base + posicoes_colunas_secundarias[4]
                                cv2.putText(imagem_alinhada, "X",
                                            (x_alternativa_E + largura_bolinha + 5, y_base + altura_bolinha - 5),
                                            cv2.FONT_HERSHEY_SIMPLEX, 1, (255, 0, 0), 2)
                        else:
                            x_inicial = x_base + posicoes_colunas_secundarias[0]
                            x_final = x_base + posicoes_colunas_secundarias[-1] + largura_bolinha
                            cv2.rectangle(imagem_alinhada, (x_inicial, y_base), (x_final, y_base + altura_bolinha),
                                          (156, 0, 156), 2)
                            cv2.putText(imagem_alinhada, "An", (x_final + 10, y_base + altura_bolinha - 5),
                                        cv2.FONT_HERSHEY_SIMPLEX, 1, (156, 0, 156), 2)
            output_file_path = os.path.join(output_folder, filename)
            cv2.imwrite(output_file_path, imagem_alinhada)

    nomes_materias = [intervalo[0] for intervalo in intervalos_nomeados]
    df_resultados = pd.DataFrame.from_dict(resultados_acertos, orient='index', columns=nomes_materias)
    df_resultados.index.name = 'Aluno'
    for materia in nomes_materias:
        total_perguntas = max_acertos_disciplina[materia]
        df_resultados[f'Nota {materia}'] = df_resultados[materia].apply(lambda acertos: f"{(acertos / total_perguntas * 10):.1f}")
    df_resultados['Total Acertos'] = df_resultados[nomes_materias].sum(axis=1)
    df_resultados['Total Erros'] = num_questoes - df_resultados['Total Acertos']
    df_resultados['Percentual de Acertos'] = (df_resultados['Total Acertos'] / num_questoes) * 100
    df_resultados['Percentual de Erros'] = 100 - df_resultados['Percentual de Acertos']
    df_resultados.reset_index(inplace=True)
    df_resultados['Aluno'] = df_resultados['Aluno'].str.replace('.png', '', case=False)

    analise = {}
    for materia in nomes_materias:
        col_nota = f'Nota {materia}'
        acima = df_resultados[df_resultados[col_nota].astype(float) >= 7].shape[0]
        abaixo = df_resultados[df_resultados[col_nota].astype(float) < 7].shape[0]
        analise[materia] = {'Acima da Média': acima, 'Abaixo da Média': abaixo}
    df_analise = pd.DataFrame(analise).T.reset_index().rename(columns={'index': 'Disciplina'})

    excel_name = f"resultados_completos_{os.path.basename(output_folder)}.xlsx"
    with pd.ExcelWriter(os.path.join(output_folder, excel_name)) as writer:
        df_resultados.to_excel(writer, sheet_name="Resultados Consolidados", index=False)
        df_analise.to_excel(writer, sheet_name="Alunos Acima-Abaixo Média", index=False)
    print(f"Resultados completos exportados para: {output_folder}/{excel_name}")

    df_acertos = df_resultados[['Aluno'] + nomes_materias]
    excel_acertos = f"acertos_por_disciplina_{os.path.basename(output_folder)}.xlsx"
    df_acertos.to_excel(os.path.join(output_folder, excel_acertos), index=False)
    print(f"Acertos por disciplina exportados para: {output_folder}/{excel_acertos}")

    colunas_notas = [f'Nota {materia}' for materia in nomes_materias]
    df_notas = df_resultados[['Aluno'] + colunas_notas]
    excel_notas = f"notas_por_disciplina_{os.path.basename(output_folder)}.xlsx"
    df_notas.to_excel(os.path.join(output_folder, excel_notas), index=False)
    print(f"Notas por disciplina exportados para: {output_folder}/{excel_notas}")

    df_totais = df_resultados[['Aluno', 'Total Acertos', 'Total Erros', 'Percentual de Acertos', 'Percentual de Erros']]
    excel_totais = f"totais_acertos_erros_{os.path.basename(output_folder)}.xlsx"
    df_totais.to_excel(os.path.join(output_folder, excel_totais), index=False)
    print(f"Totais de acertos, erros e percentuais exportados para: {output_folder}/{excel_totais}")

    for intervalo in intervalos_nomeados:
        nome_materia, inicio, fim = intervalo
        nome_materia_clean = nome_materia.replace('.png','').replace('.PNG','')
        df_disciplina = df_resultados[['Aluno', nome_materia, f'Nota {nome_materia}']].copy()
        df_disciplina.rename(columns={nome_materia: 'Acertos', f'Nota {nome_materia}': 'Nota'}, inplace=True)
        total_perguntas = max_acertos_disciplina[nome_materia]
        df_disciplina['Erros'] = total_perguntas - df_disciplina['Acertos']
        df_disciplina['Nota_num'] = df_disciplina['Nota'].astype(float)
        num_acima = df_disciplina[df_disciplina['Nota_num'] >= 7].shape[0]
        num_abaixo = df_disciplina[df_disciplina['Nota_num'] < 7].shape[0]
        df_analise_disc = pd.DataFrame({
            'Acima da Média': [num_acima],
            'Abaixo da Média': [num_abaixo]
        })

        stats = discipline_question_stats[nome_materia]
        df_questoes = pd.DataFrame({
            'Questão': list(stats.keys()),
            'Acertos': [stats[q]['acertos'] for q in stats],
            'Erros': [stats[q]['erros'] for q in stats]
        })
        df_questoes.sort_values('Questão', inplace=True)

        excel_disciplina = os.path.join(output_folder, f"{nome_materia_clean}.xlsx")
        with pd.ExcelWriter(excel_disciplina) as writer:
            df_disciplina.drop(columns=['Nota_num']).to_excel(writer, sheet_name="Desempenho dos Alunos", index=False)
            df_questoes.to_excel(writer, sheet_name="Desempenho por Questão", index=False)
            df_analise_disc.to_excel(writer, sheet_name="Acima-Abaixo Média", index=False)
        print(f"Relatório por disciplina {nome_materia_clean} exportado para: {excel_disciplina}")

    df_global = pd.DataFrame({
        'Questão': list(global_question_stats.keys()),
        'Acertos': [global_question_stats[q]['acertos'] for q in global_question_stats],
        'Erros': [global_question_stats[q]['erros'] for q in global_question_stats]
    })
    df_global.sort_values('Questão', inplace=True)
    excel_global = os.path.join(output_folder, "acertos_e_erros_gerais.xlsx")
    with pd.ExcelWriter(excel_global) as writer:
        df_global.to_excel(writer, sheet_name="Geral", index=False)
        analise_geral = {}
        for materia in nomes_materias:
            col_nota = f'Nota {materia}'
            acima = df_resultados[df_resultados[col_nota].astype(float) >= 7].shape[0]
            abaixo = df_resultados[df_resultados[col_nota].astype(float) < 7].shape[0]
            analise_geral[materia] = {'Acima da Média': acima, 'Abaixo da Média': abaixo}
        df_analise_geral = pd.DataFrame(analise_geral).T.reset_index().rename(columns={'index': 'Disciplina'})
        df_analise_geral.to_excel(writer, sheet_name="Acima-Abaixo Média", index=False)
    print(f"Relatório geral de acertos e erros exportado para: {excel_global}")

def processar_todas_as_turmas(base_dir, gabarito_base_dir, output_base_dir):
    for turma in os.listdir(base_dir):
        turma_path = os.path.join(base_dir, turma)
        output_folder = os.path.join(output_base_dir, turma)
        os.makedirs(output_folder, exist_ok=True)
        gabarito_path = os.path.join(gabarito_base_dir, f"Gabarito Oficial {turma}.png")
        print(f"Processando {turma}...")
        if turma == "1 Ano A":
            intervalos_nomeados = [
                ("FÍSICA I", 1, 4), ("FÍSICA II", 5, 8), ("QUÍMICA I", 9, 12),
                ("QUÍMICA II", 13, 16), ("BIOLOGIA I", 17, 20), ("BIOLOGIA II", 21, 24),
                ("GEOGRAFIA", 25, 28), ("HISTÓRIA GERAL", 29, 32), ("HISTÓRIA DO BRASIL", 33, 35),
                ("FILOSOFIA", 36, 38), ("SOCIOLOGIA", 39, 41),
                ("L. PORTUGUESA", 42, 46), ("LITERATURA", 47, 49),
                ("I. CIENTÍFICA", 50, 52), ("L. INGLESA", 53, 56),
                ("ED. FÍSICA", 57, 58), ("MATEMÁTICA I", 59, 63),
                ("MATEMÁTICA II", 64, 68), ("ARTE", 69, 70)
            ]
            num_questoes = 70
        elif turma == "1 Ano B":
            intervalos_nomeados = [
                ("FÍSICA I", 1, 4), ("FÍSICA II", 5, 8), ("QUÍMICA I", 9, 12),
                ("QUÍMICA II", 13, 16), ("BIOLOGIA I", 17, 20), ("BIOLOGIA II", 21, 24),
                ("GEOGRAFIA", 25, 28), ("HISTÓRIA GERAL", 29, 32), ("HISTÓRIA DO BRASIL", 33, 35),
                ("FILOSOFIA", 36, 38), ("SOCIOLOGIA", 39, 41),
                ("L. PORTUGUESA", 42, 46), ("LITERATURA", 47, 49),
                ("I. CIENTÍFICA", 50, 52), ("L. INGLESA", 53, 56),
                ("ED. FÍSICA", 57, 58), ("MATEMÁTICA I", 59, 63),
                ("MATEMÁTICA II", 64, 68), ("ARTE", 69, 70)
            ]
            num_questoes = 70
        elif turma == "2 Ano A":
            intervalos_nomeados = [
                ("FÍSICA I", 1, 5), ("FÍSICA II", 6, 10), ("QUÍMICA I", 11, 15),
                ("QUÍMICA II", 16, 20), ("BIOLOGIA I", 21, 25), ("BIOLOGIA II", 26, 30),
                ("GEOGRAFIA", 31, 34), ("HISTÓRIA GERAL", 35, 38), ("HISTÓRIA DO BRASIL", 39, 41),
                ("FILOSOFIA", 42, 44), ("SOCIOLOGIA", 45, 47),
                ("L. PORTUGUESA", 48, 52), ("LITERATURA", 53, 55),
                ("I. CIENTÍFICA", 56, 58), ("L. INGLESA", 59, 62),
                ("INT. TEXTO", 63, 67), ("MATEMÁTICA I", 66, 70),
                ("MATEMÁTICA II", 71, 75)
            ]
            num_questoes = 75
        elif turma == "2 Ano B":
            intervalos_nomeados = [
                ("FÍSICA I", 1, 5), ("FÍSICA II", 6, 10), ("QUÍMICA I", 11, 15),
                ("QUÍMICA II", 16, 20), ("BIOLOGIA I", 21, 25), ("BIOLOGIA II", 26, 30),
                ("GEOGRAFIA", 31, 34), ("HISTÓRIA GERAL", 35, 38), ("HISTÓRIA DO BRASIL", 39, 41),
                ("FILOSOFIA", 42, 44), ("SOCIOLOGIA", 45, 47),
                ("L. PORTUGUESA", 48, 52), ("LITERATURA", 53, 55),
                ("I. CIENTÍFICA", 56, 58), ("L. INGLESA", 59, 62),
                ("INT. TEXTO", 63, 67), ("MATEMÁTICA I", 66, 70),
                ("MATEMÁTICA II", 71, 75)
            ]
            num_questoes = 75
        elif turma == "3 Ano A":
            intervalos_nomeados = [
                ("FÍSICA I", 1, 5), ("FÍSICA II", 6, 10), ("QUÍMICA I", 11, 15),
                ("QUÍMICA II", 16, 20), ("BIOLOGIA I", 21, 25),
                ("BIOLOGIA II", 26, 30), ("GEOGRAFIA", 31, 34),
                ("HISTÓRIA GERAL", 35, 38), ("HISTÓRIA DO BRASIL", 39, 42),
                ("FILOSOFIA", 43, 46), ("SOCIOLOGIA", 47, 50),
                ("L. PORTUGUESA", 51, 56), ("LITERATURA", 57, 60),
                ("L. INGLESA", 61, 64), ("INT. TEXTO", 65, 68),
                ("MATEMÁTICA I", 69, 74), ("MATEMÁTICA II", 75, 80)
            ]
            num_questoes = 80
        elif turma == "3 Ano B":
            intervalos_nomeados = [
                ("FÍSICA", 1, 5), ("E.A. FÍSICA", 6, 9), ("QUÍMICA I", 10, 14),
                ("E.A. QUÍMICA", 15, 18), ("BIOLOGIA", 19, 23),
                ("E.A. BIOLOGIA", 24, 26), ("GEOGRAFIA", 27, 31),
                ("HISTÓRIA", 32, 36), ("E.A. HISTÓRIA", 37, 39),
                ("L. PORTUGUESA", 40, 45), ("E.A. L. PORTUGUESA", 46, 51),
                ("O. LITERATURA", 52, 54), ("VIDA E CARREIRA", 55, 56),
                ("BILÍNGUE", 57, 63), ("MATEMÁTICA", 64, 69),
                ("E.A. MATEMÁTICA", 70, 75), ("ARTE", 76, 78)
            ]
            num_questoes = 80

        process_files(turma_path, output_folder, gabarito_path, intervalos_nomeados, num_questoes)
base_dir = '03_pdf_to_image'
gabarito_base_dir = '03_pdf_to_image/Gabaritos Oficiais'
output_base_dir = '04_correcao_completa'

processar_todas_as_turmas(base_dir, gabarito_base_dir, output_base_dir)
print("Processamento completo.")
