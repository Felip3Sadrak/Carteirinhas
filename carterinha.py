import openpyxl
import pyautogui
import time

# Diretório padrão
diretorio = 'C:\\Users\\Benedito\\Documents\\AUTOMAÇÕES\\'

# Passo 1: Abrir a planilha Lista Piloto
lista_piloto_path = diretorio + 'ListaPilotoAluno.xlsx'
lista_piloto = openpyxl.load_workbook(lista_piloto_path)
sheet_piloto = lista_piloto.active

# Passo 2-6: Consultar e copiar dados da Lista Piloto
turma = sheet_piloto['F8'].value
periodo = sheet_piloto['D8'].value
professora = sheet_piloto['E10'].value

# Passo 3: Copiar todos os nomes da célula B12 até B39
nomes = [cell.value for row in sheet_piloto.iter_rows(min_row=12, max_row=39, min_col=2, max_col=2) for cell in row]



# Passo 7: Fechar a planilha Lista Piloto
lista_piloto.close()

# Aguardar um momento antes de continuar
time.sleep(1)

# Passo 8: Abrir a planilha Modelo da Carterinha
modelo_carterinha_path = diretorio + 'ModeloCarteirinha.xlsx'
modelo_carterinha = openpyxl.load_workbook(modelo_carterinha_path)
sheet_carterinha = modelo_carterinha.active

# Passo 9-12: Colar dados nas Carterinhas em intervalos de 14 células verticais
celula_nome = 'C8'
celula_turma = 'C9'
celula_periodo = 'C10'
celula_professora = 'E10'

for nome in nomes:
    sheet_carterinha[celula_nome].value = nome
    sheet_carterinha[celula_turma].value = turma
    sheet_carterinha[celula_periodo].value = periodo
    sheet_carterinha[celula_professora].value = professora

    # Atualizar as células para a próxima iteração
    celula_nome = celula_nome.replace('C', 'C').replace('8', str(int(celula_nome[1]) + 14))
    celula_turma = celula_turma.replace('C', 'C').replace('9', str(int(celula_turma[1]) + 14))
    celula_periodo = celula_periodo.replace('C', 'C').replace('10', str(int(celula_periodo[1]) + 14))
    celula_professora = celula_professora.replace('E', 'E').replace('10', str(int(celula_professora[1]) + 14))


# Passo 13: Salvar documento
modelo_carterinha.save(diretorio + 'carterinhas_preenchidas.xlsx')

# Aguardar um momento antes de continuar
time.sleep(1)

# Passo 14: Fechar a planilha Modelo da Carterinha
modelo_carterinha.close()
