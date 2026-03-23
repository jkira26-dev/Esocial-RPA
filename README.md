# eSocial RPA — Download Automático de XMLs
## Manual Completo de Instalação e Uso — Windows

---

## 📋 O Que Este Robô Faz

O processo é dividido em duas fases:

**FASE 1 — Criar Solicitações**
- Acessa `Download > Solicitação` do eSocial automaticamente
- Cria uma solicitação por mês, de Janeiro/2018 até a data limite do eSocial
- Salva progresso a cada solicitação (pode ser interrompido e retomado)
- Respeita o limite de 100 pedidos/dia do eSocial

**FASE 2 — Baixar os XMLs**
- Acessa `Download > Consulta` do eSocial automaticamente  
- Baixa todos os arquivos que já foram processados
- ⚠️ Execute esta fase no **dia seguinte** à Fase 1 (processamento é assíncrono)
- Arquivos ficam disponíveis por **apenas 7 dias** — não deixe passar o prazo!

---

## ⚙️ Pré-requisitos

- Windows 10 ou superior
- Python 3.10 ou superior → https://www.python.org/downloads/
  - Durante a instalação, marque: **"Add Python to PATH"**
- Google Chrome instalado
- Certificado digital A1 ou A3 configurado no Windows

---

## 🚀 Passo a Passo

### Passo 1 — Instalar

Execute o arquivo: **`1_INSTALAR.bat`**

Aguarde a instalação das dependências (Playwright). Isso é feito apenas uma vez.

---

### Passo 2 — Configurar

Abra o arquivo **`config.py`** em qualquer editor de texto (Bloco de Notas, VSCode, etc.)
e verifique/ajuste as seguintes configurações:

```python
# Data de início das solicitações
DATA_INICIO = "01/01/2018"   # ← ajuste se necessário

# Pasta onde os XMLs serão salvos
PASTA_DOWNLOAD = r"C:\esocial_xmls"   # ← ajuste para sua preferência
```

---

### Passo 3 — Abrir Chrome com Debugger

Execute o arquivo: **`2_ABRIR_CHROME.bat`**

Uma nova janela do Chrome será aberta. **Nesta janela:**
1. Acesse o eSocial normalmente
2. Selecione o certificado digital
3. Se for Procurador, informe o CNPJ e clique em Verificar
4. Aguarde chegar na tela inicial do eSocial (Menu com Empregador/Contribuinte, etc.)

⚠️ **Não feche esta janela do Chrome** — o robô precisa dela aberta!

---

### Passo 4 — Iniciar o Robô

Execute o arquivo: **`3_INICIAR_ROBO.bat`**

O menu do robô aparecerá:
```
============================================================
  eSocial RPA - Download Automático de XMLs
============================================================
  1. FASE 1 - Criar solicitações mensais
  2. FASE 2 - Baixar XMLs disponíveis
  3. Executar Fase 1 + Fase 2 em sequência
  4. Ver progresso atual
  0. Sair
============================================================
```

**Primeiro uso:** escolha `1` (Fase 1 — criar solicitações)

---

## 📅 Cronograma Recomendado

| Dia | O que fazer |
|-----|-------------|
| **Dia 1** | Execute a Fase 1 — cria todas as ~99 solicitações (~10 min) |
| **Dia 2 ou 3** | Execute a Fase 2 — baixa os XMLs prontos |
| **Atenção** | Os arquivos ficam disponíveis por **apenas 7 dias**! |

---

## ⚠️ Avisos Importantes

### Limite de 100 pedidos/dia
O eSocial permite no máximo 100 solicitações por dia. O robô já respeita esse
limite (configurado para 95 por segurança). Se houver mais de 100 meses pendentes,
o robô para automaticamente e exibe aviso. Retome no dia seguinte — o progresso
é salvo e ele continuará de onde parou.

### Sessão expira em ~15 minutos
Se o robô ficar parado por muito tempo, a sessão do eSocial pode expirar.
O robô detecta isso e pede para você fazer login novamente, sem perder o progresso.

### Arquivos disponíveis por 7 dias
Após criar as solicitações (Fase 1), os XMLs ficam disponíveis para download
por **7 dias**. Execute a Fase 2 antes desse prazo!

### Solicitações assíncronas
O processamento das solicitações pelo eSocial não é imediato. Após a Fase 1,
aguarde algumas horas antes de executar a Fase 2.

---

## 📁 Arquivos do Projeto

| Arquivo | Descrição |
|---------|-----------|
| `config.py` | ⚙️ Configurações (editar antes de usar) |
| `esocial_rpa.py` | 🤖 Script principal do robô |
| `1_INSTALAR.bat` | 🔧 Instala dependências (executar 1x) |
| `2_ABRIR_CHROME.bat` | 🌐 Abre o Chrome para login manual |
| `3_INICIAR_ROBO.bat` | ▶️ Inicia o robô |
| `progresso.json` | 📊 Criado automaticamente — não apagar! |
| `esocial_rpa.log` | 📝 Log detalhado da execução |

---

## 🔧 Solução de Problemas

**"Falha ao conectar no Chrome"**
→ Certifique-se de ter executado `2_ABRIR_CHROME.bat` antes
→ Verifique se o Chrome abriu na porta 9222 (não feche aquela janela)

**"Python não encontrado"**
→ Reinstale o Python marcando "Add Python to PATH"
→ Reinicie o computador após instalar

**"Sessão expirada"**
→ O robô vai avisar e pausar. Faça o login no Chrome e pressione ENTER

**"Solicitação não encontrou campos de data"**
→ O layout do eSocial pode ter mudado. Entre em contato para ajuste do script

**O robô travou no meio**
→ Feche e execute `3_INICIAR_ROBO.bat` novamente — ele retoma do ponto onde parou

---

## 📞 Suporte

Verifique o arquivo `esocial_rpa.log` para diagnóstico.
O arquivo `progresso.json` mostra todos os meses já processados.
