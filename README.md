# 📄 Automacao de Cadastro de Materiais – Auto Eletrica Souza

Este script automatiza o processo de cadastro de materiais no sistema SIC utilizando **Python + PyAutoGUI**. Ele le uma planilha Excel com os dados de materiais e simula a entrada manual no sistema, agrupando por **Gerador** e **Data**.

---

## ⚙️ Tecnologias Utilizadas

* Python 3.11+
* pandas
* pyautogui
* pyperclip

---

## 📂 Estrutura de Arquivos

```
automacao-cadastro/
│
├── main.py                      # Script principal da automação
├── Auto_Eletrica_Souza_Geradores.xlsm  # Planilha de entrada
└── status_cadastro.xlsx         # Planilha gerada com o status de cada grupo
```

---

## ▶️ Como Funciona

### 1. Abrir o sistema

O script inicia o sistema SIC automaticamente e realiza o login com credenciais pré-definidas.

### 2. Ler dados do Excel

A planilha contém as seguintes colunas:

* `Gerador`
* `Data`
* `ID Interno`
* `Quantidade`

As linhas são agrupadas por Gerador e Data usando `groupby`.

### 3. Preencher campos no sistema

Para cada grupo:

* Preenche o campo de Gerador
* Preenche a Data
* Cadastra item por item: Quantidade e Código do Produto
* Clica em uma posição fixa para ignorar pop-ups
* Salva o grupo
* Retorna à tela inicial

### 4. Registro de status

Gera a planilha `status_cadastro.xlsx` com os seguintes campos:

* `Gerador`
* `Data`
* `Status` → "Sim" se foi cadastrado, "Não" se falhou

---

## 📌 Trechos Importantes

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
