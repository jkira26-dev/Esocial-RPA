# ============================================================
#  CONFIGURAÇÕES DO ROBÔ - eSocial Download XML
#  Edite apenas este arquivo antes de executar.
# ============================================================

# --- DATAS ---
# Data de início das solicitações (primeiro mês a solicitar)
DATA_INICIO = "01/01/2018"

# Data limite que o eSocial exibe na tela
# (ex: "As solicitações podem ser realizadas até a data limite de 16/03/2026")
# O robô vai ler isso automaticamente da tela, mas aqui é o fallback:
DATA_LIMITE_FALLBACK = "16/03/2026"

# --- PASTA DE DOWNLOAD ---
# Pasta onde os XMLs serão salvos (Fase 2 - Download)
# Use caminho absoluto. Exemplo: r"C:\Users\jose\Downloads\esocial_xmls"
PASTA_DOWNLOAD = r"C:\esocial_xmls"

# --- BROWSER (Playwright) ---
# Se True, o robô tenta minimizar/ocultar a janela após o login bem-sucedido
BROWSER_HIDE_AFTER_LOGIN = True

# Se True, o robô roda em segundo plano desde o início (não recomendado para o eSocial)
BROWSER_HEADLESS = False

# --- COMPORTAMENTO ---
# Tempo de espera entre cada solicitação (segundos)
# Não reduzir abaixo de 3 para evitar bloqueio
PAUSA_ENTRE_SOLICITACOES = 4

# Tempo máximo de espera por carregamento de página (milissegundos)
# Aumentado pois o portal do eSocial é lento — mas agora usamos domcontentloaded
# então na prática será bem mais rápido que esse limite
TIMEOUT_PAGINA = 60000

# Tempo máximo aguardando o servidor responder após clicar em Salvar (milissegundos)
# O POST do formulário pode demorar mais que o carregamento inicial da página
TIMEOUT_POS_SALVAR = 90000

# Número máximo de tentativas por solicitação em caso de erro
MAX_TENTATIVAS = 3

# Pausa extra (segundos) quando ocorre erro antes de nova tentativa
PAUSA_APOS_ERRO = 8

# eSocial permite max 100 pedidos/dia. Se atingir o limite, pausar (horas).
LIMITE_PEDIDOS_DIA = 95  # margem de segurança

# --- ARQUIVOS DE CONTROLE ---
# Arquivo que registra quais meses já foram solicitados com sucesso
# Permite retomar o processo se interrompido
ARQUIVO_PROGRESSO = "progresso.json"

# Arquivo de log detalhado
ARQUIVO_LOG = "esocial_rpa.log"
