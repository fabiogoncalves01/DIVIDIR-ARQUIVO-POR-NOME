
# Automação de Processamento de Planilhas de Instrutores

Este projeto automatiza o processo de busca, movimentação e processamento de arquivos Excel contendo dados de horários de instrutores. O script identifica o arquivo mais recente baixado com a extensão `.xlsx`, move-o para um diretório específico, processa seus dados e cria novas planilhas separadas por instrutor.

## Funcionalidades

- **Identificação do arquivo mais recente**: O script busca o último arquivo `.xlsx` baixado no diretório de downloads.
- **Movimentação de arquivo**: O arquivo é movido para um diretório de destino específico.
- **Processamento de planilhas**: A planilha mestre é processada e os dados dos instrutores são separados em arquivos individuais.
- **Criação de novas planilhas**: Novas planilhas Excel são geradas automaticamente, uma para cada instrutor, com dados como data, horário de início, horário de fim e turma.

## Pré-requisitos

Para rodar o script, é necessário ter as seguintes bibliotecas instaladas:

- `openpyxl` (para manipulação de arquivos Excel)
- `os` (para interagir com o sistema operacional)
- `shutil` (para movimentação de arquivos)
- `re` (para limpeza e manipulação de strings)

### Instalação das dependências

Você pode instalar a biblioteca `openpyxl` via `pip`:

```bash
pip install openpyxl
```

## Como usar

1. **Configurar diretórios**:
   - Atualize as variáveis `diretorio_downloads` e `diretorio_destino` no script conforme os caminhos de diretório do seu sistema.
   
2. **Executar o script**:
   - Rode o script Python. Ele buscará automaticamente o arquivo mais recente no diretório de downloads, moverá para o diretório de destino e iniciará o processamento.
   
3. **Resultado**:
   - Após o processamento, planilhas separadas para cada instrutor serão criadas no diretório de destino.

### Exemplo de Execução

```bash
python script.py
```

Se bem-sucedido, o script exibirá o caminho do arquivo movido e criará as novas planilhas.

## Estrutura do Código

- **encontrar_ultimo_arquivo_baixado**: Identifica o arquivo mais recente com a extensão especificada no diretório de downloads.
- **mover_arquivo**: Move o arquivo identificado para um diretório de destino.
- **carregar_planilha**: Carrega a planilha principal para ser processada.
- **processar_planilha**: Processa os dados, extraindo as informações de cada instrutor e criando novas planilhas com os dados.
- **adicionar_titulos**: Adiciona os títulos das colunas nas novas planilhas criadas.
- **preencher_aba**: Preenche os dados de cada linha na aba correspondente do instrutor.


