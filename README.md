# pdfs-reader

Este projeto foi desenvolvido com o objetivo de automatizar um processo interno da área de Controladoria da Robert Bosch. A aplicação realiza a leitura de **notas de medição**, extrai automaticamente informações relevantes e as organiza em uma planilha Excel, facilitando o controle, a rastreabilidade e a eficiência na gestão de dados.

## Aviso de Confidencialidade

Este projeto **não expõe dados sigilosos, internos ou estratégicos da empresa**. 
Toda a automação depende de arquivos locais fornecidos pelo usuário, e **não há conexão com sistemas internos ou informações confidenciais da Robert Bosch**.

## Resultados Obtidos

### Redução de Tempo
  - Redução de aproximadamente **1 hora para alguns segundos** para realizar a extração e armazenamento dos dados.

### Redução de Erros
  - O processo pode ser executado por qualquer membro da equipe, mantendo consistência e padronização nos resultados.

### Confiabilidade
  - Processo pode ser executado por qualquer outra pessoa da equipe, mantendo os mesmos padrões e resultados.

## Funcionalidades

- **Identificação e download automático das notas de medições:** A aplicação identifica as contas conectadas, localiza a caixa de entrada correta e realiza o download automático dos PDFs, salvando-os tanto em Documents quanto em um caminho de rede predefinido.
- **Extração de dados (`extract_data`)**: Realiza a leitura dos PDFs das notas de medição, extraindo informações específicas a partir de uma pasta selecionada pelo usuário.
- **Armazenamento dos dados**: As informações extraídas são organizadas e registradas em uma planilha Excel, salva localmente e também em uma pasta compartilhada na rede.


## Tecnologias

- Python 3.x
- Flet Desktop
- Bibliotecas padrão e adicionais via `pip`

**Todas as dependências estão listadas no arquivo `requirements.txt`.**

## Como executar o Projeto

```bash
# 1. Clone o repositório
git clone https://github.com/fernandaatim/pdfs-reader.git
cd pdfs-reader

# 2. (Opcional) Crie um ambiente virtual
python -m venv venv

# 3. Ative o ambiente virtual
# Windows
venv\Scripts\activate
# Linux/macOS
source venv/bin/activate

# 4. Instale as dependências
pip install -r requirements.txt

# 5. Execute a aplicação
py src/ui/ui.py
```

## Como gerar o Executável

Execute o seguinte comando no terminal:
```
flet pack src/ui/ui.py --name "pdf-reader" --add-data "src/ui/assets/;assets" --add-data "src/modules;modules" --icon "src/ui/assets/icons/icon_bosch.ico" --hidden-import pdfplumber --hidden-import openpyxl --hidden-import flet --hidden-import pythoncom --hidden-import win32com --hidden-import win32timezone
```

## Licença

Este projeto é de **uso interno**, e **não deve ser comercializado ou redistribuído** sem autorização prévia.
Todos os direitos relacionados à marca **Robert Bosch** são reservados à própria empresa.
