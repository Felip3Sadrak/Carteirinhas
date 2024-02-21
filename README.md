# Automação de Geração de Carterinhas com Python e Excel

Este projeto visa automatizar o processo de geração de carterinhas estudantis a partir de dados fornecidos em planilhas do Excel. A automação é implementada em Python, utilizando a biblioteca `openpyxl` para manipulação de planilhas.

## Funcionalidades

1. **Seleção da Pasta de Automações:**
   - Ao executar a aplicação, o usuário é solicitado a selecionar a pasta onde estão localizados os arquivos necessários para a geração das carteirinhas.
   - Isso proporciona flexibilidade ao usuário para escolher a localização dos arquivos de entrada.

2. **Verificação de Arquivos:**
   - Após a seleção da pasta, a aplicação verifica automaticamente se os arquivos "Lista Piloto-2024.xlsx" e "ModeloCarteirinha.xlsx" estão presentes na pasta selecionada.
   - Se algum dos arquivos estiver ausente, a aplicação exibirá uma mensagem de erro e encerrará o processo.

2. **Geração de Carterinhas:**
   - Abre a planilha "Modelo da Carterinha" que serve como modelo para as carterinhas.
   - Cola os nomes nas carterinhas a partir da célula C8, em intervalos de 14 células verticais para cada novo nome.
   - Cola as informações de Turma, Período e Nome da Professora em intervalos correspondentes.
   - 
3. **Feedback ao Usuário:**
   - Ao final do processo de geração, a aplicação exibe uma mensagem de sucesso, informando ao usuário que as carteirinhas foram geradas com êxito.
   - 
4. **Salvamento e Fechamento:**
   - Salva as carterinhas geradas em um novo arquivo.
   - Fecha a planilha "Modelo da Carterinha".

## Pré-requisitos

- Python 3.x
- Bibliotecas Python: `openpyxl`

## Como Usar

1. Clone o repositório para sua máquina local.
2. Instale as dependências necessárias utilizando `pip install -r requirements.txt`.
3. Execute o script `carterinha.py`.

