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

# Verificar a existência do arquivo
if not os.path.exists(modelo_carterinha_path):
    raise FileNotFoundError("Certifique-se de que o arquivo ModeloCarteinha.xlsx existe.")

# Abrir planilha Modelo Carteirinha
modelo_carterinha = openpyxl.load_workbook(modelo_carterinha_path)
sheet_carterinha = modelo_carterinha.active

# Inicializar variáveis de posição
linha_inicial_c12 = 12
linha_inicial_j12 = 12

# Iterar sobre os dados copiados em ciclos de 20 nomes
for i in range(0, len(nomes), 20):
    # Determinar qual coluna utilizar (C8 ou J8)
    coluna = 3 if i // 20 % 2 == 0 else 10
    linha = linha_inicial_c12 if coluna == 3 else linha_inicial_j12

    # Atribuir valor à célula correspondente na planilha Modelo Carteirinha
    for j, nome in enumerate(nomes[i:i + 20]):
        sheet_carterinha.cell(row=linha + j * 14, column=coluna, value=nome)

    # Verificar se atingiu a célula C274 ou J274
    if i + 20 < len(nomes):
        # Se houver mais nomes, copiar os próximos 20
        dados_sheet['A1'] = 'Nomes'
        dados_sheet.append(nomes[i + 20:i + 40])
     
# Salvar e fechar planilha Modelo Carteirinha
nome_turma = turma.replace(" ", " ")  # Substituir espaços por underscores, se necessário
nome_arquivo_carteirinhas = f'Carteirinhas-{nome_turma}.xlsx'
modelo_carterinha.save(os.path.join(pasta_automacoes, nome_arquivo_carteirinhas))
modelo_carterinha.close()

# Excluir a planilha Dados
os.remove(os.path.join(pasta_automacoes, 'Dados.xlsx'))