# 🛠️ Automação de Cadastro de Materiais - Auto Elétrica Souza

Este projeto é uma solução completa de **RPA (Automação de Processos Robóticos)** integrada entre **Excel VBA**, **Python**, **Power Query** e **Power BI**, desenvolvida para agilizar o processo de cadastro de materiais utilizados em manutenções de geradores de energia no sistema SIC da empresa **Auto Elétrica Souza**.

---

## 📚 Contexto do Projeto

Antes dessa automação, o processo de cadastro de materiais era **manual**, repetitivo e sujeito a erros. Para resolver isso, foi criado um fluxo inteligente que envolve:

- Cadastro de materiais via **formulário VBA no Excel**
- Exportação dos dados
- Automação do cadastro no SIC com **Python + PyAutoGUI**
- Geração de logs de execução
- Consolidação de dados via **Power Query**
- Análise de custos e manutenção via **Power BI**

---

## 🧱 Estrutura Completa do Processo

### 1. Cadastro via Formulário VBA (Excel)

- Usuário preenche os campos obrigatórios:
  - **Gerador**
  - **Data**
  - **Técnico**
  - **Materiais**
  - **ID Externo**
  - **ID Interno**
  - **Quantidade**
- Cada linha cadastrada recebe um **ID Tabela incremental** automaticamente.

**Validações implementadas no VBA:**

- Campo obrigatório para cada informação
- Verificação de datas válidas
- Preenchimento automático dos IDs ao selecionar o material
- Ordenação automática após cada cadastro
- Possibilidade de excluir registros pelo **formulário "ExcluirID"**

---

### 2. Exportação da Base de Dados

- A base de dados fica armazenada na aba "**Cadastro de materiais**".
- As colunas utilizadas são:

| Coluna | Descrição |
|---|---|
| F | Gerador |
| G | Data |
| H | Técnico |
| I | Materiais |
| J | ID Externo |
| K | ID Interno |
| L | Quantidade |
| M | ID Tabela |

---

### 3. Automação com Python (PyAutoGUI)

O script `main.py` faz a leitura da planilha e interage automaticamente com o sistema SIC, clicando e preenchendo os campos como se fosse um operador humano.

#### Tecnologias usadas no Python:

- **pandas**
- **openpyxl**
- **pyautogui**
- **pyperclip**
- **tkinter**

#### Fluxo da Automação Python:

1. Abre o sistema SIC.
2. Realiza login.
3. Navega até o módulo de cadastro de saída de materiais.
4. Faz a leitura dos dados agrupados por **Gerador** e **Data**.
5. Realiza o preenchimento automático de cada item.
6. Trata pop-ups e confirmações de "Sim/Não".
7. Gera um log final com o status de cada linha (Sucesso ou Falha).

#### Estrutura de pastas:

```plaintext
automacao-souza/
├── imagens/            # Prints dos botões e telas do SIC para o PyAutoGUI
├── main.py             # Código Python da automação
├── logs_automacao/     # Onde os logs em Excel são salvos após cada execução
```

### 4. Consolidação dos Logs via Power Query

- Todos os logs de execução (gerados pelo Python) ficam na pasta `/logs_automacao`.
- No **Power Query**, foi criado um processo de **importação em lote** desses logs.
- Cada vez que um novo log é gerado, o Power Query carrega automaticamente na consolidação geral.

---

### 5. Cálculo de Custos de Manutenção

O processo de cálculo de custos envolve:

1. **Exportação da tabela de produtos do SIC**, contendo os **IDs dos materiais e seus valores unitários**.
2. Uso de uma planilha Excel para realizar um **ÍNDICE + CORRESP**, cruzando os IDs dos materiais do log com os valores extraídos do SIC.
3. Geração de uma **tabela dinâmica** no Excel para calcular o **custo total de cada manutenção**, separada por gerador, data e técnico responsável.
4. Esta tabela consolidada serve de base para o **Dashboard no Power BI**.

---

### 6. Integração com o Dashboard Power BI

O Dashboard Power BI consome os dados da planilha consolidada para entregar os seguintes indicadores:

| Métrica | Exemplo |
|---|---|
| Quantidade de Manutenções | Por período |
| Custo Total | Por gerador, por técnico |
| Materiais mais utilizados | Por categoria |
| Evolução de Gastos | Por mês |

- Layout com segmentações por **gerador**, **data**, **técnico** e **tipo de material**.
- Visualizações com **gráficos de barras**, **linhas** e **mapas de calor** para identificar os maiores pontos de custo.

---
# 🚀 Primeiros Passos e Como Rodar

Para utilizar e manter este projeto, siga os passos abaixo:

---

## ⚙️ Pré-requisitos

- Microsoft Excel (2016 ou superior)  
- Python 3.9+ (recomendado, por conta do uso de `os.startfile` e `datetime`)  
- Power BI Desktop  
- Acesso ao sistema SIC com credenciais válidas  

---

## 📦 Instalação das Dependências (Python)

1. Navegue até a pasta `automacao-souza/` no terminal.  
2. Crie um arquivo `requirements.txt` na raiz do projeto com o conteúdo:
- pandas
- pyautogui
- pyperclip
- openpyxl

3. Instale as dependências com:

```bash
pip install -r requirements.txt
ou
pip install pandas pyautogui pyperclip openpyxl
```
## ▶️ Como Executar a Automação

1. Ajuste os caminhos no `main.py`:

   - Atalho do SIC:  
     `C:\Users\Lucas\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\sic.lnk`

   - Pasta para salvar logs:  
     `C:\Users\Lucas\OneDrive\Trabalho\Planilhas de excel\logs_automacao`

   > **Atenção:** Ajuste esses caminhos para o ambiente da sua máquina.

2. Abra o arquivo Excel com o formulário VBA, ex: `Cadastro_Materiais_AE_Souza.xlsm`.

3. Cadastre os materiais via formulário VBA na aba **"Cadastro de materiais"** (recomendado).

4. Salve a planilha.

5. Certifique-se de que os técnicos deram baixa nos materiais usados.

6. Execute a automação Python:

```bash
python main.py
```
Na interface que abrir:

- Clique em **Selecionar Planilha** e escolha seu arquivo Excel.  
- Clique em **Iniciar Automação** para rodar o processo.  
- Aguarde o término e confira a mensagem de sucesso.  
- Verifique os logs gerados na pasta configurada (`logs_automacao`).  

---

## 📊 Visualização no Power BI

- Abra o arquivo Power BI, ex: `Dashboard_Custos_Manutencao.pbix`.  
- Certifique-se que o Power Query atualizou os dados dos logs.  
- Caso necessário, clique em **Atualizar** para carregar os dados mais recentes.  
- O Power BI está configurado para atualizar automaticamente todos os dias às 10h via Power Automate (não é preciso atualizar manualmente todo dia).  

---

## 💡 Fluxo Recomendado Resumido

- Cadastre os materiais na planilha principal usando o formulário VBA (recomendado).  
- Salve o arquivo.  
- Após os técnicos darem baixa nos materiais usados, rode a automação Python conforme tutorial acima.  
- Confira o log para garantir que tudo ocorreu bem.  
- Atualize a planilha que consolida os logs via Power Query.  
- Analise o custo da manutenção.  
- Alimente essa informação na planilha principal, que serve de base para o Power BI.  


### 🎯 Diferenciais Técnicos do Projeto

- Integração real entre **VBA**, **Python**, **Power Query** e **Power BI**.
- Uso de **interface Tkinter** para facilitar o uso da automação por qualquer colaborador.
- Geração de **logs Excel** para rastreamento completo.
- Tratamento de exceções tanto no **VBA** quanto no **Python**.
- Organização de pastas seguindo boas práticas de versionamento e manutenção.
- Processo **100% reprodutível** e **documentado**.

---

### ✅ Melhorias planejadas (Backlog)

- Implementar **OCR por texto (Tesseract)** para eliminar a dependência de imagens nos cliques do PyAutoGUI.
- Criar um **instalador .exe** para facilitar a instalação da automação em outros computadores da empresa.
- Melhorar a granularidade dos logs, registrando erros linha a linha com mais detalhes.
- Integrar a base de dados com um **banco relacional (ex: SQLite ou PostgreSQL)**.
- Criar um **painel web** (ex: com Streamlit ou Flask) para acompanhamento das execuções e logs em tempo real.
- Implementar **testes automatizados** no Python.

---

### ⚠️ Importante

Este projeto é **privado** e de **uso exclusivo da Auto Elétrica Souza**.

O código, imagens, planilhas e demais ativos não devem ser compartilhados fora da empresa sem autorização expressa.

---

### 🤝 Autor

Desenvolvido por: **Lucas Amorim**

📧 Email: lucas.amorim.porciuncula@gmail.com 
🔗 LinkedIn: https://www.linkedin.com/in/lucas-amorim-powerbi/

---


