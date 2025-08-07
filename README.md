# BOT_LendoValoresPDF

Este projeto é um script Python que lê uma pasta de arquivos PDF, extrai informações específicas de cada PDF e salva os dados em uma planilha Excel na aba "Valores_PDF".

## Funcionalidades

- Varre uma pasta com arquivos PDF
- Identifica a existência de uma aba especifica
- Realiza validações no conteúdo do PDF e no nome do arquivo.
- Extrai valores específicos de cada PDF, como "Valor 1" (sendo esse em R$) e "Valor 2"(sendo em kWh).
- Salva os dados extraídos em uma aba específica de um arquivo Excel.
- Gera mensagens de erro detalhadas para arquivos que não seguem o padrão esperado.

## Como usar

1. Instale as dependências do projeto:
   ```bash
   pip install -r requirements.txt
   ```

2. Altere algumas informações iniciais para o que fizer mais sentido para você e seu código
   - Caso sua moeda não seja R$ (Real Brasileiro) altere na função `def limpar_numero(valor)`, os primeiros argumentos do método `replace` para os da sua moeda, caso não seja Real
   - Mude o nome da aba `Valores PDF`para um que faça sentido para você em `# Seleciona a aba "Valores_PDF"`
   - Mude o cabeçalho para um que faça sentido para você em `# Cria o cabeçalho`
   - Caso tenha mais variáveis a serem encontradas, precisa colocar todas como None em `# Inicializa as variáveis como None a cada iteração`
   - Mude o número da página em `# MUDE AQUI O NÚMERO DA PÁGINA SE FOR NECESSÁRIO`


3. Adapte o código para validar os padrões que você identifique em todos os PDFs
   - Pode ser um título como colocado em `# VALIDAÇÃO DE TÍTULO DENTRO DO PDF`
   - Pode ser um uma data que mude todo mês como colocado em `# VALIDAÇÃO DE PERÍODO DENTRO DO PDF`
      -> Deve se sinalizar ao editor que mude a data todo mês (Ex: `------------ MUDAR TODO MÊS -------------`)
      -> Deve se lembrar de alterar todo o mês (Ao perceber que todos darão erro, será lembrado)
   - Pode ser o nome do arquivo como em `# VALIDAÇÃO DE NOME DO ARQUIVO UTILIZANDO UM NOME PADRÃO A TODOS` através das técnicas de regex.search que estão explicadas abaixo do passo a passo.
   - Pode ser uma validação entre o nome do arquivo com o nome interno do arquivo, assim verificando se o documento está nomeado no padrão certo mas com o nome errado. Ex: `# VALIDAÇÃO DE NOME DO ARQUIVO DENTRO DO PDF`

4. Agora o código vai buscar os valores que deseja encontrar, fique atento pois ao modificar uma parte, precisará modificar tudo.
   - O código exemplo usa um método de loop por cada linha até encontrar parte de um texto ("Valor 1:" ou "Valor 2:")
   - Após identificar a linha correta, ele usa o método regex.search para retirar um valor que obedece a um certo padrão.
   - Caso não encontre, a variavel continuará sendo None
   - Depois a variável é limpada usando a função `def limpar_numeros(valor)`
   - Usamos 2 métodos para verificar se as variáveis estão no que podemos chamar de exceções, enviando mensagens e alertas para quando caem nesses casos. EX: `valor_1 = R$ 2,07` ou `valor_2 = 0`
   - Esses valores são colocados no excel em `aba.append` para futuramente usarmos para análises

5. Tratamento de erros:
   - Os erros de execução podem ser tratados na forma de `try:... except:` assim continuando o loop pelos PDFs
   - Os erros que não impedem a execução, como valores zerados, arquivos nomeados errados, podem ser tratados da forma `if.... continue` assim pulando para o próximo item do loop e indicando o erro através do `aba.append` (para mostrar no excel) ou `print("ERRO")` (para mostrar no terminal)

6. No script `BOT_LendoValoresPDF.py`, altere as variáveis:
   - `caminho_pasta`: Caminho para a pasta contendo os PDFs.
   - `caminho_excel`: Caminho para o arquivo Excel onde os dados serão salvos.

7. Execute o script:
   ```bash
   python BOT_LendoValoresPDF.py
   ```

## Explicação detalhada dos comandos e estruturas do código

Abaixo, cada comando, função, método, estrutura ou chamada utilizada no código é explicada. A ordem segue o script, e não há repetições.

   ### `import`
   - Importa módulos externos ou internos para serem usados no código.

   ### `from ... import ...`
   - Importa partes específicas de um módulo (ex: uma função ou classe) para acesso direto.

   ### `def nome_da_funcao(...):`
   - Define uma função, agrupando código para ser reutilizado quantas vezes precisar.

   ### `return`
   - Finaliza a execução de uma função e retorna um valor para quem chamou.

   ### `float(valor)`
   - Converte uma string ou número para o tipo float (ponto flutuante).

   ### `[... for variavel in lista if ...]`
   - Estrutura de repetição que percorre elementos de um objeto (lista, string, numeros) e coloca todos dentro de uma lista se corresponderem a condição if

   ### `arquivo.name`
   - Retorna apenas o nome do arquivo, sem o caminho completo.

   ### `Path(pasta)`
   - Cria um objeto de caminho (Path) para o diretório especificado, facilita os erros comuns com barras invertidas, caminhos relativos, etc.

   ### `Path.iterdir()`
   - Lista todos os arquivos e pastas dentro do diretório especificado.

   ### `arquivo.is_file()`
   - Retorna True se o objeto (arquivo) for realmente um arquivo (e não uma pasta).

   ### `if ...:  else:`
   - Estrutura condicional. Executa o bloco `if` se a condição for verdadeira e o bloco `else` se for falsa (pode não ter else).

   ### `replace(old, new)`
   - Substitui partes de uma string por outra (ex: vírgula por ponto).

   ### `try: ... except: ...`
   - Bloco para capturar e tratar erros durante a execução de um trecho de código. Tenta realizar o código, caso não consiga por algum erro de execução, pula para o except.

   ### `None`
   - Representa a ausência de valor em Python.

   ### `openpyxl.load_workbook(caminho_excel)`
   - Abre um arquivo Excel existente e salva numa variável.

   ### `try: ... except: ...`
   - Bloco para capturar e tratar erros durante a execução de um trecho de código.

   ### `wb.sheetnames`
   - Lista os nomes das abas (sheets) do Excel.

   ### `wb["Nome_Aba"]`
   - Acessa uma aba específica pelo nome.

   ### `wb.remove(aba)`
   - Remove uma aba do arquivo Excel.

   ### `wb.create_sheet("Nome_Aba")`
   - Cria uma nova aba no arquivo Excel com o "Nome_Aba".

   ### `aba.append([valores])`
   - Adiciona uma nova linha ao final da aba do Excel.

   ### `for variavel in lista:`
   - Estrutura de repetição que percorre elementos de uma lista ou sequência.

   ### `caminho_pasta / arquivo`
   - Cria um caminho de arquivo combinando diretório e nome de arquivo usando Path.

   ### `with ... as ...:`
   - Segura um uso de um recurso para botar realizar diversas funções nele.

   ### `pdfplumber.open(caminho_arquivo)`
   - Abre um arquivo PDF para leitura com a biblioteca pdfplumber.

   ### `pdf.pages[0]`
   - Acessa a primeira página do PDF.

   ### `extract_text()`
   - Extrai texto de uma página de PDF.

   ### `if objeto is None:`
   - Testa se o objeto está vazio ("None").

   ### `print()`
   - Imprime mensagens no terminal.

   ### `f"texto qualquer {variavel}"`
   - O f anterior a uma string possibilita a entrada de variáveis no meio da string por meio de {}

   ### `continue`
   - Pula para a próxima iteração do loop, sem executar o restante do bloco atual.

   ### `not in`
   - Retorna True se um elemento NÃO estiver em uma sequência (inverso de "in").

   ### `split("\n")`
   - Divide uma string em várias partes, usando a quebra de linha ("\n") como separador.

   ### `re.search(padrao, texto)`
   - Procura um padrão usando expressão regular no texto, retorna Match se encontrar. Muitas vezes usa caracteres e simbolos regex para facilitar a busca.

      #### `\s`
      - **Significado:** Corresponde a qualquer caractere de espaço em branco (espaço, tabulação, quebra de linha).
      - **Exemplo:** `\s+` — um ou mais espaços em branco.

      #### `+`
      - **Significado:** Corresponde a uma ou mais repetições do elemento anterior.
      - **Exemplo:** `a+` — "a", "aa", "aaa", etc.

      #### `()`
      - **Significado:** Agrupa parte da expressão para captura ou aplicar quantificadores.
      - **Exemplo:** `(abc)+` — "abc", "abcabc", etc.

      #### `.`
      - **Significado:** Corresponde a qualquer caractere, exceto uma nova linha.
      - **Exemplo:** `a.b` — corresponde a "acb", "a9b", "a b", etc.

      #### `*`
      - **Significado:** Corresponde a zero ou mais repetições do elemento anterior.
      - **Exemplo:** `a*` — "", "a", "aa", "aaa", etc.

      #### `?`
      - **Significado:** Torna o elemento anterior opcional (zero ou uma vez).
      - **Exemplo:** `a?` — "", "a"

      #### `[...]`
      - **Significado:** Define um conjunto de caracteres. Corresponde a qualquer um dos caracteres dentro dos colchetes.
      - **Exemplo:** `[abc]` — "a", "b" ou "c"

      #### `\d`
      - **Significado:** Corresponde a qualquer dígito numérico (0–9).
      - **Exemplo:** `\d{2}` — exatamente dois dígitos.

      #### Exemplos do código
      - `r"Reembolso\s+(.*?)\s+- Mar 25"`  
         - `Reembolso\s+` — "Reembolso" seguido de um ou mais espaços
         - `(.*?)` — captura qualquer coisa (de forma não gananciosa) entre "Reembolso" e " - Mar 25"
         - `\s+- Mar 25` — espaço(s) seguidos de "- Mar 25"

      - `r"R\$\s*([\d,.]+)"`  
         - `R\$` — procura o caractere "R$"
         - `\s*` — zero ou mais espaços em branco
         - `([\d,.]+)` — captura um grupo com um ou mais dígitos, vírgulas ou pontos

      - `r"([\d,.]+)\s*kWh"`  
         - `([\d,.]+)` — captura um grupo com um ou mais dígitos, vírgulas ou pontos
         - `\s*` — zero ou mais espaços em branco
         - `kWh` — procura o texto literal "kWh"

   ### `match.group(n)`
   - Acessa o grupo n capturado pela expressão regular.

   ### `elif ...:`
   - Bloco condicional alternativo a um if, executado se o if anterior falhar e a condição do elif for verdadeira.

   ### `except Exception as e:`
   - Captura a exceção e armazena o erro na variável e.

   ### `str(e)`
   - Converte o erro capturado em texto para exibição ou registro.

   ### `wb.save(caminho_excel)`
   - Salva as alterações feitas no arquivo Excel.

   ### `r""` (string raw)
   - Indica uma string em formato bruto (raw), onde barras invertidas não são interpretadas como caracteres especiais (útil para caminhos de arquivos).


## Observações

- Certifique-se de ajustar os filtros de período e padrões de nome de arquivo conforme o mês/período de interesse.
- O script sobrescreve a aba "Valores_PDF" a cada execução.

## Licença

Este projeto está sob a licença MIT.