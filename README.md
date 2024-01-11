# Automação de Geração de Carterinhas com Python e Excel

Este projeto visa automatizar o processo de geração de carterinhas estudantis a partir de dados fornecidos em planilhas do Excel. A automação é implementada em Python, utilizando a biblioteca `openpyxl` para manipulação de planilhas e `pyautogui` para interação com a interface gráfica.

## Funcionalidades

1. **Abertura e Consulta de Dados:**
   - Abre a planilha "Lista Piloto" e consulta a primeira planilha de turma.
   - Copia todos os nomes presentes entre as células B12 e B39.
   - Copia informações como Turma, Período e Nome da Professora.

2. **Geração de Carterinhas:**
   - Abre a planilha "Modelo da Carterinha" que serve como modelo para as carterinhas.
   - Cola os nomes nas carterinhas a partir da célula C8, em intervalos de 14 células verticais para cada novo nome.
   - Cola as informações de Turma, Período e Nome da Professora em intervalos correspondentes.

3. **Salvamento e Fechamento:**
   - Salva as carterinhas geradas em um novo arquivo.
   - Fecha a planilha "Modelo da Carterinha".

## Pré-requisitos

- Python 3.x
- Bibliotecas Python: `openpyxl`, `pyautogui`

## Como Usar

1. Clone o repositório para sua máquina local.
2. Instale as dependências necessárias utilizando `pip install -r requirements.txt`.
3. Execute o script `carterinha.py`.

