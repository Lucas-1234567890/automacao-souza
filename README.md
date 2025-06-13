# ⚙️ Automação de Cadastro de Materiais – Auto Elétrica Souza

Este projeto automatiza o processo de cadastro de materiais no sistema SIC usando **Python** e **PyAutoGUI**, simulando entradas humanas a partir de dados em uma planilha Excel. Cada material é inserido conforme agrupamentos por **Gerador** e **Data**.

---

## 🧰 Tecnologias Utilizadas

- Python 3.11+
- pandas
- pyautogui
- pyperclip
- openpyxl
- tkinter (para pop-ups visuais)

---

## 📁 Estrutura de Arquivos

```
automacao-cadastro/
│
├── main.py # Script principal da automação
├── imagens/ # Prints usados para localizar elementos na tela
│ ├── atualizacao.png
│ ├── saida.png
│ └── ...
├── Auto_Eletrica_Souza_Geradores.xlsm # Planilha de entrada
└── log_resultado_automacao.xlsx # Planilha de log gerada com o status final
```

---


---

## ▶️ Como Funciona

### 1. Inicialização

- Exibe um pop-up de boas-vindas com `tkinter.messagebox`
- Abre o sistema SIC automaticamente via atalho `.lnk`
- Faz login usando credenciais definidas no código

### 2. Leitura da Planilha

A planilha precisa conter as colunas:

- `Gerador`
- `Data`
- `ID Interno`
- `Quantidade`

As linhas são agrupadas por `Gerador` e `Data` para processar em blocos.

### 3. Preenchimento no Sistema

Para cada grupo:

- Preenche o campo Gerador
- Preenche a Data
- Insere cada material: Quantidade e Código
- Clica em posição fixa (1010, 617) para fechar pop-ups inesperados
- Salva os dados e volta para tela inicial

### 4. Log e Exportação

- Adiciona uma coluna `Status` para indicar "Sim"/"Não"
- Registra `Data Registro` da automação
- Junta com log anterior (se existir) sem sobrescrever
- Salva tudo no `log_resultado_automacao.xlsx`
- Exibe pop-up final com caminho do arquivo salvo

---

## 📌 Trechos-Chave

### 🖼️ Localização de Elementos via Imagem

```python
caminho_imagem = os.path.join("imagens", imagem)
pyautogui.locateOnScreen(caminho_imagem, grayscale=True, confidence=0.8)

### Função de localização com timeout

```python
def encontrar_imagem(imagem):
    ...
```

Procura uma imagem na tela por até 20 segundos.

---

### Agrupamento da planilha

```python
grupos = tabela.groupby(["Gerador", "Data"])
```

Agrupa os dados por Gerador e Data para cadastro em blocos.

---

### Loop principal de cadastro

```python
for (gerador, data), grupo in grupos:
```

Percorre os dados por grupo e cadastra item a item com tratamento de erro e registro de status.

---

### Extras adicionados

* ✅ Pop-ups ignorados via clique fixo `(1010, 617)`
* ✅ Redução de `sleep()` para acelerar preenchimento
* ✅ Planilha `status_cadastro.xlsx` com resultados
* ✅ Pop-up visual no início e no fim com `pyautogui.alert()`


# Explicação do uso de `groupby` com loop `for` em pandas

O método `groupby` do pandas agrupa os dados de um DataFrame com base em uma ou mais colunas. O resultado é um objeto iterável que retorna pares de chave e grupo.

## Como funciona:

```python
grupos = tabela.groupby(["Coluna1", "Coluna2"])

for (valor_col1, valor_col2), grupo in grupos:
    print(f"Grupo: ({valor_col1}, {valor_col2})")
    print(grupo)
    print("-" * 40)

Grupo: (G1, 2024-06-10)
  Gerador        Data  Quantidade
0      G1  2024-06-10           5
1      G1  2024-06-10           3
----------------------------------------
Grupo: (G2, 2024-06-11)
  Gerador        Data  Quantidade
2      G2  2024-06-11           7
----------------------------------------
Grupo: (G2, 2024-06-12)
  Gerador        Data  Quantidade
3      G2  2024-06-12           2
----------------------------------------


---

## 🚀 Como Rodar

1. Instale os pacotes:

```bash
pip install pyautogui pandas pyperclip openpyxl
```

2. Execute o script:

```bash
python main.py
```

3. Verifique o arquivo `status_cadastro.xlsx` ao final.

---

## 🔐 Segurança

* Repositório privado por conter interações com sistema interno.
* Senhas de acesso devem ser mantidas seguras fora do script.

---

## 👨‍💼 Autor

**Lucas Amorim**
Auxiliar Administrativo • Estudante de Engenharia de Dados & IA
