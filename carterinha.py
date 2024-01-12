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

# Iterar sobre os dados copiados
for i, nome in enumerate(nomes):
    # Encontrar a célula correspondente na planilha Modelo Carteirinha
    # Suponha que você tenha os índices da linha e coluna da célula mesclada
    linha_mesclada = 8  # ajuste conforme necessário
    coluna_mesclada = 3  # ajuste conforme necessário
    
    # Calcular a linha principal da célula mesclada
    linha_principal = linha_mesclada + i * 14
    
    # Atribuir valor à célula principal
    sheet_carterinha.cell(row=linha_principal, column=coluna_mesclada, value=nome)

# Salvar e fechar planilha Modelo Carteirinha
modelo_carterinha.save('C:/Users/Benedito/Documents/GitHub/Carteirinhas/CarterinhasPreenchidas.xlsx')
modelo_carterinha.close()
dados_workbook.close()
