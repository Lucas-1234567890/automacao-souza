# 🤖 Sistema de Automação - Auto Elétrica Souza

> **Sistema integrado de automação para gestão de materiais, relatórios e comunicação via WhatsApp**

Este repositório contém uma suíte de scripts Python desenvolvidos para automatizar processos operacionais da Auto Elétrica Souza, incluindo:

- ✅ **Cadastro automatizado no sistema SIC**
- 📊 **Geração de relatórios de geradores**
- 📱 **Envio automático via WhatsApp**
- 🔧 **Controle de manutenção preventiva**
- ⌨️ **Sistema de atalhos de teclado**

---

## 📋 Pré-requisitos

### Software Necessário
- **Python 3.8+** ([Download aqui](https://www.python.org/downloads/))
- **Google Chrome** (para automação WhatsApp)
- **Microsoft Excel** (para manipular planilhas)
- **Sistema SIC** instalado e configurado

### Conhecimentos Básicos
- Uso básico do terminal/prompt de comando
- Conceitos básicos de Excel (planilhas, abas, colunas)
- WhatsApp Web configurado no navegador

---

## 🚀 Instalação

### 1. Clone o Repositório
```bash
git clone https://github.com/seu-usuario/automacao-souza.git
cd automacao-souza
```

### 2. Crie um Ambiente Virtual
```bash
# Windows
python -m venv .venv
.venv\Scripts\activate

# Linux/Mac  
python3 -m venv .venv
source .venv/bin/activate
```

### 3. Instale as Dependências
```bash
pip install -r requirements.txt
```

**Principais bibliotecas utilizadas:**
- `pandas` - Manipulação de dados Excel
- `selenium` - Automação web (WhatsApp)
- `pyautogui` - Automação desktop (SIC)
- `tkinter` - Interface gráfica
- `ttkbootstrap` - Interface moderna
- `openpyxl` - Leitura/escrita Excel

### 4. Configure o Ambiente
1. **Crie as pastas necessárias:**
```bash
mkdir imagens
mkdir logs
```

2. **Configure os caminhos no arquivo `main.py`:**
   - Caminho do executável SIC
   - Pasta de logs
   - Planilhas Excel

---

## 📖 Uso

### 🎯 Execução Rápida com Atalhos

**Para usuários que querem começar imediatamente:**

1. **Execute o arquivo `.bat` fornecido:**
   ```bash
   # Duplo clique no arquivo ou execute via terminal:
   automacao-souza.bat
   ```

2. **Use os atalhos de teclado registrados:**
   - `Ctrl + Alt + W` → Envio de WhatsApp
   - `Ctrl + Alt + P` → Manutenção Preventiva  
   - `Ctrl + Alt + A` → Automação SIC
   - `Ctrl + Alt + E` → Relatórios de Entrada/Saída

---

### 📱 WhatsApp - Envio de Relatórios (`whatsapp.py`)

**Função:** Envia relatórios personalizados de uso de materiais para técnicos via WhatsApp.

```bash
python whatsapp.py
```

**O que faz:**
1. Lê a planilha `Auto_Elétrica_Souza_Geradores.xlsm`
2. Agrupa dados por técnico, gerador e data
3. Gera mensagens personalizadas
4. Envia via WhatsApp Web automaticamente
5. Mantém log de envios para evitar duplicatas

**Configurações importantes:**
```python
# Modifique estas variáveis no início do arquivo:
CAMINHO_ARQUIVO = r"C:\caminho\para\sua\planilha.xlsm"
ABA = "Cadastro de materiais"
CAMINHO_LOG = r"C:\caminho\para\logs\log_envios.csv"
```

---

### 🔧 Manutenção Preventiva (`preventiva.py`)

**Função:** Envia relatório de status das manutenções preventivas de geradores.

```bash
python preventiva.py
```

**Características:**
- ✅ Status visual com emojis (Em dia/Vencido/Hoje/Amanhã)
- 📅 Leitura automática de datas
- 🎯 Envio para lista de responsáveis
- ⏰ Saudação inteligente baseada no horário

**Para configurar novos telefones:**
```python
# Edite a lista no arquivo preventiva.py:
TELEFONES = ["71999365938", "71988776655"]  # Adicione novos números
```

---

### 🤖 Automação SIC (`main.py`)

**Função:** Interface gráfica completa para automação do sistema SIC com cadastro em lote.

```bash
python main.py
```

**Funcionalidades:**
- 🖥️ Interface gráfica moderna com abas
- 📊 Preview dos dados antes da execução
- ⚡ Processamento em lote com estatísticas em tempo real
- 📝 Sistema de logs detalhado
- ⚙️ Configurações personalizáveis
- 🛑 Controle de parada de emergência

**Fluxo de uso:**
1. Selecione a planilha Excel
2. Verifique o preview dos dados na aba correspondente
3. Configure caminhos na aba "Configurações"
4. Execute a automação
5. Acompanhe o progresso em tempo real

---

### 📊 Relatórios de Geradores (`entrada_saida.py`)

**Função:** Sistema para criação e envio de relatórios detalhados de geradores.

```bash
python entrada_saida.py
```

**Interface inclui:**
- 📝 Formulário de cadastro de geradores
- 👁️ Preview das mensagens antes do envio  
- 🗂️ Lista organizadas dos relatórios
- 📱 Envio direto para WhatsApp
- 🔄 Validação de campos obrigatórios

---

### ⌨️ Sistema de Atalhos (`atalhos.py`)

**Função:** Registra atalhos globais de teclado para execução rápida dos scripts.

```bash
python atalhos.py
```

**Atalhos disponíveis:**
| Tecla | Script | Função |
|-------|--------|---------|
| `Ctrl+Alt+W` | whatsapp.py | Envio WhatsApp |
| `Ctrl+Alt+P` | preventiva.py | Manutenção |
| `Ctrl+Alt+A` | main.py | Automação SIC |
| `Ctrl+Alt+E` | entrada_saida.py | Relatórios |

> **💡 Dica:** Mantenha este script rodando em segundo plano para usar os atalhos a qualquer momento.

---

## 📁 Estrutura do Repositório

```
automacao-souza-main/
├── 📄 main.py                    # Interface principal - Automação SIC
├── 📱 whatsapp.py               # Envio automatizado WhatsApp  
├── 🔧 preventiva.py             # Relatórios de manutenção
├── 📊 entrada_saida.py          # Sistema de relatórios de geradores
├── ⌨️ atalhos.py                # Gerenciador de atalhos de teclado
├── 🚀 automacao-souza.bat       # Script de inicialização rápida
├── 📋 requirements.txt          # Dependências Python
├── 📖 README.md                 # Este arquivo
├── 📁 imagens/                  # Screenshots para automação PyAutoGUI
│   ├── atualizacao.png
│   ├── saida.png
│   ├── gerador.png
│   └── ... (outras imagens)
├── 📁 logs/                     # Logs e histórico de execuções
└── 📁 .venv/                    # Ambiente virtual Python
```

### 🗂️ Arquivos Importantes

| Arquivo | Propósito | Quando Modificar |
|---------|-----------|------------------|
| **main.py** | Sistema principal com interface gráfica | Nunca modificar diretamente |
| **config_automacao.json** | Configurações salvas automaticamente | Gerado automaticamente |
| **imagens/*.png** | Screenshots para automação | Quando interface do SIC mudar |
| **requirements.txt** | Lista de dependências | Ao adicionar novas bibliotecas |

---

## ⚙️ Configuração Avançada

### 🖼️ Atualizando Screenshots (Imagens)

Se a interface do sistema SIC mudar, você precisará atualizar as imagens:

1. **Capture nova screenshot:**
   - Use a ferramenta de captura do Windows (`Win + Shift + S`)
   - Salve como PNG na pasta `imagens/`
   - **Importante:** Mantenha o mesmo nome do arquivo

2. **Teste a precisão:**
```python
import pyautogui
# Teste se a imagem é encontrada
resultado = pyautogui.locateOnScreen('imagens/sua_imagem.png', confidence=0.8)
print(f"Imagem encontrada: {resultado is not None}")
```

### 📊 Configurando Novas Planilhas

Para trabalhar com planilhas diferentes:

```python
# Em whatsapp.py, modifique:
CAMINHO_ARQUIVO = r"C:\nova\planilha.xlsm"
ABA = "Nome da nova aba"

# Ajuste as colunas se necessário:
tabela = pd.read_excel(CAMINHO_ARQUIVO, sheet_name=ABA, 
                       skiprows=3, usecols="F:N")  # Ajuste usecols
```

### 🔧 Ajustes de Performance

**Para computadores mais lentos:**
```python
# Aumente os tempos de espera em main.py:
sleep(2)  # Aumente para sleep(3) ou sleep(4)

# Reduza a confiança de reconhecimento de imagem:
confidence = 0.7  # Era 0.8, reduza para 0.7 ou 0.6
```

**Para computadores mais rápidos:**
```python
# Diminua os tempos de espera:
sleep(0.5)  # Reduza para sleep(0.3)
```

---

## 🛠️ Boas Práticas e Recomendações

### ✅ Antes de Modificar Qualquer Script

1. **Faça backup:**
```bash
cp script_original.py script_original_backup.py
```

2. **Teste em ambiente separado**
3. **Documente mudanças realizadas**

### 🔒 Segurança e Dados Sensíveis

- **Nunca commite** números de telefone reais
- **Use variáveis de ambiente** para caminhos sensíveis:
```python
import os
CAMINHO_PLANILHA = os.getenv('PLANILHA_PATH', 'caminho_default.xlsm')
```

### 📈 Monitoramento e Logs

**Sempre verifique os logs após execução:**
- Logs da automação SIC: pasta `logs/`
- Logs do WhatsApp: `log_envios.csv`
- Logs de erro: console/terminal

**Para debug detalhado:**
```python
# Adicione estas linhas no início dos scripts:
import logging
logging.basicConfig(level=logging.DEBUG)
```

### 🚫 Problemas Comuns e Soluções

| Problema | Possível Causa | Solução |
|----------|----------------|---------|
| "Módulo não encontrado" | Ambiente virtual não ativado | Execute `.venv\Scripts\activate` |
| "Imagem não encontrada" | Interface SIC mudou | Atualize screenshots |
| "Erro de permissão" | Arquivo Excel aberto | Feche o Excel antes de executar |
| "WhatsApp não abre" | ChromeDriver desatualizado | Execute `pip install --upgrade webdriver-manager` |

### 🔄 Atualizações e Manutenção

**Atualize dependências regularmente:**
```bash
pip install --upgrade -r requirements.txt
```

**Verifique compatibilidade:**
- Tente executar todos os scripts após atualizações
- Mantenha backup das versões funcionais

---

## 📞 Contato e Suporte

### 👨‍💻 Desenvolvedor Principal
**Lucas Amorim Porciuncula** - Engenharia de Dados e IA  
- 📧 Email: [lucas.amorim.porciuncula@gmail.com]
- 💼 LinkedIn: [(https://www.linkedin.com/in/lucas-amorim-powerbi/)]
- 🐛 Issues: [Crie uma issue neste repositório](https://github.com/seu-usuario/automacao-souza/issues)

### 🆘 Em Caso de Problemas

1. **Primeiro, verifique os logs** de erro no terminal
2. **Consulte a seção "Problemas Comuns"** acima  
3. **Crie uma issue detalhada** com:
   - Descrição do erro
   - Captura de tela
   - Log de erro completo
   - Passos para reproduzir

### 📚 Documentação Adicional

- 📖 [Documentação Python](https://docs.python.org/3/)
- 🐼 [Pandas Docs](https://pandas.pydata.org/docs/)
- 🌐 [Selenium Docs](https://selenium-python.readthedocs.io/)
- 🤖 [PyAutoGUI Docs](https://pyautogui.readthedocs.io/)

---

## 📜 Histórico de Versões

| Versão | Data | Alterações |
|--------|------|------------|
| **2.0** | 2024-12 | Interface gráfica completa, sistema de logs |
| **1.5** | 2024-11 | Adicionado sistema de atalhos |
| **1.0** | 2024-10 | Versão inicial com scripts básicos |

---

## 📄 Licença

Este projeto é de uso interno da **Auto Elétrica Souza**. Todos os direitos reservados.

> **⚠️ Importante:** Este software foi desenvolvido especificamente para os processos internos da empresa. Não redistribuir sem autorização.

---

**🎯 Feito com dedicação para otimizar processos e aumentar a produtividade da equipe!**
