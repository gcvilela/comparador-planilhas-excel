# Comparador de Planilhas Excel - Coincidências e Diferenças

Este script tem como objetivo comparar duas planilhas do Excel (`1.xlsx` e `2.xlsx`), identificando **coincidências** e **diferenças** com base em colunas específicas, como **SENHA**, **PACIENTE**, **USUÁRIO** e **NÚMERO GUIA DA OPERADORA**.

## 🛠️ Funcionalidades

- Normaliza e padroniza os textos (remove acentos, espaços e coloca tudo em maiúsculas).
- Gera chaves compostas para realizar comparações precisas entre registros.
- Identifica:
  - Registros presentes em ambas as planilhas (`Coincidências Lado a Lado`)
  - Registros exclusivos de cada planilha
  - Coincidências alternativas com base em SENHA vs. NÚMERO GUIA
- Exporta os resultados em um único arquivo Excel com múltiplas abas.

## 📂 Estrutura Esperada

O script espera encontrar dois arquivos no **mesmo diretório onde o executável é executado**:

- `1.xlsx` — Planilha com colunas: `PACIENTE`, `SENHA`
- `2.xlsx` — Planilha com colunas: `USUÁRIO`, `NÚMERO GUIA DA OPERADORA`

## 📤 Saída

O programa gera o arquivo:

- `resultado_comparacao.xlsx`

Com as seguintes abas:

1. **Coincidências Lado a Lado** — Registros presentes nas duas planilhas com mesma chave composta.
2. **Não Coincidências 1** — Registros da planilha 1 que não estão na planilha 2.
3. **Não Coincidências 2** — Registros da planilha 2 que não estão na planilha 1.
4. **Coincidências Senha-Guia** — Registros que coincidem ao comparar SENHA com NÚMERO GUIA, e PACIENTE com USUÁRIO.

## ▶️ Como Usar

### Opção 1: Executando via Python

1. Instale as dependências (caso necessário):

   ```bash
   pip install pandas openpyxl
   ```

2. Coloque os arquivos `1.xlsx` e `2.xlsx` no mesmo diretório do script.
3. Execute o script:

   ```bash
   python comparador.py
   ```

### Opção 2: Executável

Se você transformou o script em um executável (por exemplo, usando `PyInstaller`):

1. Coloque os arquivos `1.xlsx` e `2.xlsx` no mesmo diretório do `.exe`.
2. Execute o programa com um duplo clique ou via terminal.
3. O arquivo `resultado_comparacao.xlsx` será gerado no mesmo local.

## 🧠 Lógica de Comparação

- As colunas de interesse são **normalizadas**: acentos e caracteres especiais são removidos, espaços são eliminados e tudo é convertido para maiúsculas.
- Chaves compostas são geradas:
  - **Planilha 1:** `SENHA | PACIENTE`
  - **Planilha 2:** `NÚMERO GUIA | USUÁRIO`
- Depois, são feitas comparações diretas entre as chaves e também entre os valores brutos de `SENHA` e `NÚMERO GUIA`.

## ⚠️ Requisitos

- Python 3.7 ou superior
- Bibliotecas:
  - `pandas`
  - `openpyxl`
  - `unicodedata` (builtin)

## 🧰 Como gerar o executável (opcional)

Caso queira gerar um `.exe` com [PyInstaller](https://pyinstaller.org/):

```bash
pip install pyinstaller
pyinstaller --onefile comparador.py
```

O executável será criado dentro da pasta `dist`.

---

## 📄 Licença

Este projeto é livre para uso e modificação. Adapte conforme necessário para suas necessidades.