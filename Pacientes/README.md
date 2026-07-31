# Consulta de Pacientes

Aplicativo desktop simples em Python que importa dados de pacientes de uma planilha Excel para um banco SQLite local e disponibiliza uma interface gráfica (Tkinter) para busca por nome.

## Funcionalidades

- Leitura automática de uma planilha `.xlsx` com os dados dos pacientes.
- Importação/atualização dos dados em um banco SQLite local (`pacientes.db`).
- Interface gráfica com campo de busca por nome (filtro em tempo real, conforme você digita).
- Tabela de resultados com rolagem e contador de pacientes encontrados.
- Seletor de arquivo integrado, caso a planilha padrão não seja encontrada.

## Colunas utilizadas da planilha

O script procura pelas seguintes colunas no cabeçalho (primeira linha) da planilha:

| Coluna            |
|-------------------|
| Atendimento       |
| Data Atendimento  |
| Prontuário        |
| Nome              |
| Data Registro     |

> A planilha pode conter outras colunas (ex.: Código CID, Diagnóstico, Idade, Sexo, Convênio, Plano) — elas são ignoradas pela importação, que utiliza apenas as cinco colunas acima.

## Requisitos

- Python 3.8+
- [openpyxl](https://pypi.org/project/openpyxl/)
- Tkinter (já incluso na instalação padrão do Python no Windows; no Linux, instale se necessário)

Instalação das dependências:

```bash
pip install openpyxl
```

No Linux, se o Tkinter não estiver disponível:

```bash
sudo apt install python3-tk
```

## Como usar

Coloque o arquivo `Relacao_Pacientes_Organizado.xlsx` na mesma pasta do script (ou informe o caminho manualmente) e execute:

```bash
python pacientes_app.py
```

Ou, informando o caminho da planilha diretamente:

```bash
python pacientes_app.py "C:\caminho\para\Relacao_Pacientes_Organizado.xlsx"
```

Se o arquivo padrão não for encontrado, uma janela para seleção manual do arquivo será aberta automaticamente.

Ao rodar o script:

1. A planilha é lida e os dados são gravados no banco `pacientes.db` (a tabela é recriada a cada execução, evitando duplicidade).
2. A janela de consulta é aberta, permitindo buscar pacientes pelo nome.

## Estrutura do projeto

```
.
├── pacientes_app.py                    # Script principal (importação + interface gráfica)
├── Relacao_Pacientes_Organizado.xlsx   # Planilha de origem dos dados
└── pacientes.db                        # Banco SQLite gerado/atualizado automaticamente
```

## Observações

- Toda vez que o script é executado, a tabela `pacientes` no banco é limpa e repopulada com os dados atuais da planilha, garantindo que o banco sempre reflita a última versão do arquivo `.xlsx`.
- As datas são convertidas para o formato `DD/MM/AAAA` na hora da importação.
- A busca por nome não diferencia maiúsculas/minúsculas e localiza qualquer trecho do nome digitado.

## Licença

Uso interno / livre para adaptação conforme a necessidade do projeto.
