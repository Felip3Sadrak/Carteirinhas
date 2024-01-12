import openpyxl
import os

# Caminhos para as planilhas
pasta_automacoes = 'C:/Users/Benedito/Documents/GitHub/Carteirinhas'
lista_piloto_path = os.path.join(pasta_automacoes, 'ListaPilotoAluno.xlsx')
modelo_carterinha_path = os.path.join(pasta_automacoes, 'ModeloCarteirinha.xlsx')

# Verificar a existência do arquivo
if not os.path.exists(lista_piloto_path):
    raise FileNotFoundError("Certifique-se de que o arquivo ListaPilotoAluno.xlsx existe.")

# Abrir a planilha Lista Piloto
lista_piloto = openpyxl.load_workbook(lista_piloto_path)
sheet_piloto = lista_piloto.active

# Obter os nomes da lista
nomes = [cell.value for row in sheet_piloto.iter_rows(min_row=12, max_row=39, min_col=2, max_col=2) for cell in row]

# Obter os dados da turma, período e nome da professora
turma = sheet_piloto['F8'].value
periodo = sheet_piloto['D8'].value
nome_professora = sheet_piloto['E10'].value

# Criar uma nova planilha chamada DADOS
dados_workbook = openpyxl.Workbook()
dados_sheet = dados_workbook.active

# Preencher a nova planilha com as informações
dados_sheet['A1'] = 'Nomes'
dados_sheet.append(nomes)
dados_sheet['A8'] = 'Turma'
dados_sheet['B8'] = turma
dados_sheet['A9'] = 'Período'
dados_sheet['B9'] = periodo
dados_sheet['A10'] = 'Nome da Professora'
dados_sheet['B10'] = nome_professora

# Salvar a nova planilha
dados_workbook.save(os.path.join(pasta_automacoes, 'Dados.xlsx'))

# Fechar planilha Lista piloto
lista_piloto.close()

# Abrir planilha Modelo Carteirinha
modelo_carterinha = openpyxl.load_workbook(modelo_carterinha_path)
sheet_carterinha = modelo_carterinha.active

# Inicializar variáveis de posição
linha_inicial_c8 = 8
linha_inicial_j8 = 8

# Iterar sobre os dados copiados
for i, nome in enumerate(nomes):
    # Determinar qual coluna e linha utilizar (C8 ou J8)
    if i == 0:
        coluna, linha = 3, linha_inicial_c8
    elif i == 261:
        coluna, linha = 3, 274
    else:
        coluna, linha = 10, linha_inicial_j8

    # Atribuir valor à célula correspondente na planilha Modelo Carteirinha
    sheet_carterinha.cell(row=linha + i * 14, column=coluna, value=nome)

# Salvar e fechar planilha Modelo Carteirinha
modelo_carterinha.save(os.path.join(pasta_automacoes, 'CarterinhasPreenchidas.xlsx'))
modelo_carterinha.close()