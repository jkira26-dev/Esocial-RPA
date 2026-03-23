"""
==============================================================
  eSocial RPA - Download Automático de XMLs
  Desenvolvido para: Windows + Python + Playwright
  
  FASE 1: Cria todas as solicitações mensais (Jan/2018 em diante)
  FASE 2: Baixa os XMLs das solicitações prontas
==============================================================
"""

import asyncio
import json
import logging
import os
import re
import sys
from calendar import monthrange
from datetime import date, datetime, timedelta
from pathlib import Path

from playwright.async_api import async_playwright, TimeoutError as PlaywrightTimeout

# Importa configurações
sys.path.insert(0, str(Path(__file__).parent))
import config

# ─── Estatísticas da Sessão ──────────────────────────────────────────────────

class SessionStats:
    """Coleta métricas durante a execução para gerar o relatório final."""

    def __init__(self):
        self.inicio = datetime.now()
        self.fim = None
        self.opcao_escolhida = ""
        self.data_inicio_fase1 = ""
        self.data_limite = ""

        # Fase 1
        self.f1_total_meses = 0
        self.f1_ja_no_progresso = 0
        self.f1_novos_sucesso = 0
        self.f1_ja_existia = 0
        self.f1_falhas: list[str] = []          # períodos que falharam
        self.f1_tentativas_extras = 0           # retries usados
        self.f1_sessoes_expiradas = 0
        self.f1_limite_diario_atingido = False
        self.f1_interrompido_por_erros = False

        # Fase 2
        self.f2_disponiveis = 0
        self.f2_baixados = 0
        self.f2_ja_baixados = 0
        self.f2_erros: list[str] = []

    def encerrar(self):
        self.fim = datetime.now()

    @property
    def duracao(self) -> str:
        if not self.fim:
            return "—"
        delta = self.fim - self.inicio
        h, rem = divmod(int(delta.total_seconds()), 3600)
        m, s = divmod(rem, 60)
        return f"{h:02d}h {m:02d}m {s:02d}s"

    @property
    def taxa_sucesso_f1(self) -> str:
        processados = self.f1_novos_sucesso + self.f1_ja_existia + len(self.f1_falhas)
        if processados == 0:
            return "N/A"
        taxa = ((self.f1_novos_sucesso + self.f1_ja_existia) / processados) * 100
        return f"{taxa:.1f}%"


# ─── Configuração de Log ─────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(config.ARQUIVO_LOG, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger("esocial_rpa")


# ─── Controle de Progresso ───────────────────────────────────────────────────

def _arquivo_progresso(cnpj: str = None) -> str:
    """Retorna o caminho do arquivo de progresso (global ou por CNPJ)."""
    if cnpj:
        cnpj_limpo = re.sub(r'\D', '', str(cnpj))
        return str(Path(config.ARQUIVO_PROGRESSO).parent / f"progresso_{cnpj_limpo}.json")
    return config.ARQUIVO_PROGRESSO


def carregar_progresso(cnpj: str = None) -> dict:
    """Carrega o arquivo de progresso. Se cnpj fornecido, usa arquivo por empresa."""
    arquivo = _arquivo_progresso(cnpj)
    if Path(arquivo).exists():
        with open(arquivo, "r", encoding="utf-8") as f:
            return json.load(f)
    return {"solicitacoes_criadas": [], "downloads_concluidos": []}


def salvar_progresso(progresso: dict, cnpj: str = None):
    """Salva o progresso. Se cnpj fornecido, usa arquivo por empresa."""
    arquivo = _arquivo_progresso(cnpj)
    with open(arquivo, "w", encoding="utf-8") as f:
        json.dump(progresso, f, indent=2, ensure_ascii=False)


# ─── Handler de Log para GUI ─────────────────────────────────────────────────

class CallbackLogHandler(logging.Handler):
    """Envia registros de log para um callback (usado pela GUI)."""
    def __init__(self, callback):
        super().__init__()
        self.callback = callback
        self.setFormatter(logging.Formatter("%(message)s"))

    def emit(self, record):
        try:
            self.callback({
                "tipo": "log",
                "level": record.levelname.lower(),
                "text": self.format(record),
                "ts": datetime.now().strftime("%H:%M:%S"),
            })
        except Exception:
            pass


# ─── Utilidades de Data ──────────────────────────────────────────────────────

def ultimo_dia_mes(ano: int, mes: int) -> int:
    """Retorna o último dia do mês."""
    return monthrange(ano, mes)[1]


def gerar_meses(data_inicio_str: str, data_limite_str: str):
    """
    Gera lista de tuplas (data_inicio, data_fim) para cada mês
    entre data_inicio e data_limite.
    Formato esperado: DD/MM/YYYY
    """
    def parse(s):
        return datetime.strptime(s.strip(), "%d/%m/%Y").date()

    inicio = parse(data_inicio_str)
    limite = parse(data_limite_str)

    meses = []
    ano, mes = inicio.year, inicio.month
    while True:
        primeiro = date(ano, mes, 1)
        ultimo = date(ano, mes, ultimo_dia_mes(ano, mes))
        if primeiro > limite:
            break
        # Se o último dia do mês ultrapassar o limite, usar o limite
        if ultimo > limite:
            ultimo = limite
        meses.append((
            primeiro.strftime("%d/%m/%Y"),
            ultimo.strftime("%d/%m/%Y"),
        ))
        # Avança para o próximo mês
        if mes == 12:
            mes = 1
            ano += 1
        else:
            mes += 1
    return meses


# ─── Conexão com Chrome ──────────────────────────────────────────────────────

async def conectar_chrome(playwright, silencioso: bool = False):
    """
    Conecta ao Chrome já aberto com certificado digital.
    O usuário precisa ter iniciado o Chrome com --remote-debugging-port=9222
    """
    if not silencioso:
        log.info(f"Conectando ao Chrome em {config.CHROME_DEBUG_URL} ...")
    try:
        browser = await playwright.chromium.connect_over_cdp(config.CHROME_DEBUG_URL)
        log.info("✅ Conectado ao Chrome com sucesso!")
        return browser
    except Exception as e:
        log.error(f"❌ Falha ao conectar no Chrome: {e}")
        log.error('  Execute: 2_ABRIR_CHROME.bat e faça login no eSocial.')
        raise


async def obter_aba_esocial(browser):
    """Retorna a aba do eSocial já aberta ou abre uma nova."""
    contextos = browser.contexts
    if not contextos:
        raise RuntimeError("browser_fechado")

    context = contextos[0]
    paginas = context.pages

    # Verifica se as páginas existentes ainda estão vivas
    for page in paginas:
        try:
            if "esocial.gov.br" in page.url:
                # Testa se a página realmente responde
                await page.evaluate("1+1")
                log.info(f"✅ Aba do eSocial encontrada: {page.url}")
                return page
        except Exception:
            continue  # Página fechada/morta, tenta a próxima

    # Nenhuma aba válida — abre uma nova
    log.info("Nenhuma aba ativa do eSocial. Abrindo nova aba...")
    try:
        page = await context.new_page()
        await page.goto(
            "https://www.esocial.gov.br/portal/Home/Inicial",
            timeout=config.TIMEOUT_PAGINA,
            wait_until="domcontentloaded",
        )
        return page
    except Exception:
        raise RuntimeError("browser_fechado")


async def aguardar_reconexao(playwright) -> tuple:
    """
    Detectou que o Chrome foi fechado.
    Pausa e aguarda o usuário reabrir o Chrome e fazer login.
    Retorna (browser, page) após reconexão bem-sucedida.
    """
    log.error("🔴 Chrome foi fechado ou perdeu conexão!")
    print()
    print("  ╔══════════════════════════════════════════════════════╗")
    print("  ║  ⚠️  CHROME FECHADO — AÇÃO NECESSÁRIA               ║")
    print("  ╠══════════════════════════════════════════════════════╣")
    print("  ║  1. Execute novamente:  2_ABRIR_CHROME.bat           ║")
    print("  ║  2. Faça login no eSocial com o certificado          ║")
    print("  ║  3. Aguarde a tela inicial do eSocial carregar       ║")
    print("  ║  4. Volte aqui e pressione ENTER                     ║")
    print("  ╚══════════════════════════════════════════════════════╝")
    print()

    while True:
        input("  Pressione ENTER após fazer o login no eSocial...")
        try:
            browser = await conectar_chrome(playwright, silencioso=True)
            page = await obter_aba_esocial(browser)
            log.info("✅ Reconexão bem-sucedida! Retomando execução...")
            return browser, page
        except Exception as e:
            print(f"  ❌ Ainda não foi possível conectar: {e}")
            print("  Certifique-se de ter aberto o Chrome com 2_ABRIR_CHROME.bat")
            print()



# ─── FASE 1: Criar Solicitações ──────────────────────────────────────────────

async def ler_data_limite_tela(page) -> str:
    """
    Tenta ler a data limite que o eSocial exibe na tela de solicitação.
    Ex: "As solicitações podem ser realizadas até a data limite de 16/03/2026."
    """
    try:
        await page.goto(
            "https://www.esocial.gov.br/portal/download/Pedido/Solicitacao",
            timeout=config.TIMEOUT_PAGINA,
            wait_until="domcontentloaded",
        )
        # Aguarda apenas o elemento que nos interessa, não scripts externos
        await page.wait_for_selector(".alert-info", timeout=10000)

        texto = await page.inner_text("body")
        match = re.search(r"data limite de\s+(\d{2}/\d{2}/\d{4})", texto)
        if match:
            data_limite = match.group(1)
            log.info(f"📅 Data limite lida da tela: {data_limite}")
            return data_limite
    except Exception as e:
        log.warning(f"Não foi possível ler a data limite da tela: {e}")

    log.info(f"📅 Usando data limite do config: {config.DATA_LIMITE_FALLBACK}")
    return config.DATA_LIMITE_FALLBACK


async def preencher_data(page, campo_id: str, valor: str):
    """
    Preenche um campo de data do eSocial pelo ID.
    Usa JavaScript para garantir que o valor seja aceito mesmo com máscaras/datepicker.
    """
    # 1. Clica no campo para focar
    await page.click(f"#{campo_id}")
    await page.wait_for_timeout(300)

    # 2. Seleciona todo o conteúdo atual via Ctrl+A
    await page.keyboard.press("Control+a")
    await page.wait_for_timeout(100)

    # 3. Injeta o valor diretamente via JavaScript (evita conflito com máscara)
    await page.evaluate(
        """([id, val]) => {
            var el = document.getElementById(id);
            el.value = val;
            // Dispara eventos para que o framework Angular/jQuery reconheça a mudança
            el.dispatchEvent(new Event('input', { bubbles: true }));
            el.dispatchEvent(new Event('change', { bubbles: true }));
            el.dispatchEvent(new Event('blur', { bubbles: true }));
        }""",
        [campo_id, valor]
    )
    await page.wait_for_timeout(200)


async def criar_solicitacao(page, data_inicio: str, data_fim: str, tentativa: int = 1) -> bool:
    """
    Cria uma solicitação de download para o período especificado.
    Usa os IDs exatos do HTML do eSocial.
    Retorna True se bem-sucedido.
    """
    log.info(f"  📋 Criando solicitação: {data_inicio} → {data_fim} (tentativa {tentativa})")

    try:
        # 1. Navegar para a tela de Nova Solicitação
        #    domcontentloaded: dispara assim que o HTML está pronto, sem esperar
        #    scripts de analytics, chatbot e outros recursos externos do portal.
        await page.goto(
            "https://www.esocial.gov.br/portal/download/Pedido/Solicitacao",
            timeout=config.TIMEOUT_PAGINA,
            wait_until="domcontentloaded",
        )
        # Aguarda apenas o formulário — ignora o resto da página
        await page.wait_for_selector("#TipoPedido", state="visible", timeout=15000)

        # 2. Verificar se a sessão ainda está ativa
        conteudo = await page.inner_text("body")
        if "selecione o seu perfil" in conteudo.lower() or "sessão expirou" in conteudo.lower():
            log.error("❌ SESSÃO EXPIRADA! O usuário precisa fazer login novamente.")
            raise RuntimeError("Sessão expirada")

        # 3. Selecionar o dropdown "Tipo de solicitação" (id="TipoPedido")
        log.debug("    Selecionando tipo de solicitação...")
        await page.select_option("#TipoPedido", value="1")
        await page.wait_for_timeout(600)

        # 4. Aguardar os campos de data aparecerem (o JS remove a classe "hide")
        log.debug("    Aguardando campos de data ficarem visíveis...")
        await page.wait_for_selector("#DataInicial", state="visible", timeout=8000)

        # 5. Preencher Data de Início
        log.debug(f"    Preenchendo DataInicial: {data_inicio}")
        await preencher_data(page, "DataInicial", data_inicio)

        # 6. Preencher Data de Fim
        log.debug(f"    Preenchendo DataFinal: {data_fim}")
        await preencher_data(page, "DataFinal", data_fim)

        # 7. Fechar calendário popup se estiver aberto
        await page.click("h2.titulo-tabela")
        await page.wait_for_timeout(300)

        # 8. Clicar em Salvar via JavaScript — evita o timeout de navegação do Playwright
        #    O portal faz um POST normal (não SPA), então a página recarrega.
        #    Usar JS para submeter evita ficar preso esperando a navegação terminar.
        log.debug("    Submetendo formulário...")
        await page.wait_for_selector("#btnSalvar", state="visible", timeout=8000)

        # Submete e aguarda o DOM estar pronto — sem esperar networkidle
        async with page.expect_navigation(
            wait_until="domcontentloaded",
            timeout=config.TIMEOUT_POS_SALVAR,
        ):
            await page.evaluate("document.getElementById('btnSalvar').click()")

        # 9. Aguarda o #mensagemGeral aparecer (sucesso ou erro do servidor)
        try:
            await page.wait_for_selector(
                "#mensagemGeral .alert, #mensagemGeral .retornoServidor",
                timeout=8000,
            )
        except PlaywrightTimeout:
            pass  # Pode não aparecer se redirecionou direto — trata abaixo

        # 10. Lê a mensagem retornada no #mensagemGeral (sucesso ou erro)
        msg_geral = ""
        try:
            msg_geral = await page.inner_text("#mensagemGeral")
            msg_geral = msg_geral.strip()
        except:
            pass

        conteudo_pos = await page.inner_text("body")

        # --- Caso 1: Solicitação criada com sucesso ---
        if "enviada com sucesso" in conteudo_pos.lower():
            log.info(f"  ✅ Criado com sucesso: {data_inicio} → {data_fim}")
            return "sucesso"

        # --- Caso 2: Já existe pedido para esse período ---
        # Mensagem exata do eSocial: "Pedido não foi aceito. Já existe um pedido do mesmo tipo."
        if "já existe um pedido" in msg_geral.lower() or "já existe um pedido" in conteudo_pos.lower():
            log.info(f"  ⏭️ Já existe pedido para {data_inicio} → {data_fim} — pulando.")
            return "ja_existe"

        # --- Caso 3: Redirecionou para Consulta (alguns casos de sucesso) ---
        if "consulta" in page.url.lower():
            log.info(f"  ✅ Redirecionado para Consulta — sucesso: {data_inicio} → {data_fim}")
            return "sucesso"

        # --- Caso 4: Algum outro erro do servidor ---
        if msg_geral:
            log.error(f"  ❌ Mensagem do servidor: {msg_geral}")
        else:
            log.error(f"  ❌ Resposta inesperada para {data_inicio} → {data_fim}")
        return "falha"

    except RuntimeError:
        raise  # Propaga: sessão expirada ou browser_fechado
    except PlaywrightTimeout as e:
        log.error(f"  ❌ Timeout: {e}")
        return "falha"
    except Exception as e:
        msg = str(e).lower()
        if "closed" in msg or "target" in msg and "closed" in msg:
            # Chrome foi fechado — precisa reconectar
            raise RuntimeError("browser_fechado")
        log.error(f"  ❌ Erro inesperado: {e}")
        return "falha"


def perguntar_data_inicio(padrao: str) -> str:
    """
    Pergunta ao usuário a partir de qual mês deseja iniciar as solicitações.
    Aceita MM/AAAA ou DD/MM/AAAA. Retorna sempre no formato DD/MM/AAAA (dia 1).
    Pressionar ENTER sem digitar nada usa o valor padrão do config.py.
    """
    print()
    print("  ┌─────────────────────────────────────────────────────┐")
    print("  │         A PARTIR DE QUAL MÊS INICIAR?               │")
    print("  ├─────────────────────────────────────────────────────┤")
    print(f"  │  Padrão (config.py): {padrao:<32}│")
    print("  │  Formatos aceitos  : MM/AAAA  ou  DD/MM/AAAA        │")
    print("  │  Exemplos          : 02/2024  ou  01/02/2024         │")
    print("  │  [ENTER]           : usa o padrão do config.py       │")
    print("  └─────────────────────────────────────────────────────┘")

    while True:
        entrada = input("  Digite o mês/ano de início: ").strip()

        # ENTER sem digitar → usa o padrão
        if not entrada:
            print(f"  ✅ Usando padrão: {padrao}")
            log.info(f"Data de início: {padrao} (padrão do config)")
            return padrao

        # Formato MM/AAAA → converte para 01/MM/AAAA
        if re.fullmatch(r"\d{2}/\d{4}", entrada):
            mes, ano = entrada.split("/")
            data_formatada = f"01/{mes}/{ano}"

        # Formato DD/MM/AAAA → usa direto
        elif re.fullmatch(r"\d{2}/\d{2}/\d{4}", entrada):
            data_formatada = entrada

        else:
            print("  ❌ Formato inválido. Use MM/AAAA (ex: 02/2024) ou DD/MM/AAAA (ex: 01/02/2024)")
            continue

        # Valida se é uma data real
        try:
            datetime.strptime(data_formatada, "%d/%m/%Y")
            print(f"  ✅ Início definido: {data_formatada}")
            log.info(f"Data de início escolhida: {data_formatada}")
            return data_formatada
        except ValueError:
            print(f"  ❌ Data inválida: {data_formatada}. Verifique dia e mês.")


async def fase1_criar_solicitacoes(page, progresso: dict, stats: SessionStats,
                                    playwright=None, callback=None,
                                    cnpj: str = None,
                                    data_inicio_override: str = None):
    """FASE 1: Itera mês a mês criando as solicitações."""

    # Instala o handler de log para GUI se callback fornecido
    cb_handler = None
    if callback:
        cb_handler = CallbackLogHandler(callback)
        log.addHandler(cb_handler)

    try:
        log.info("=" * 60)
        log.info("  FASE 1 — CRIANDO SOLICITAÇÕES MENSAIS")
        log.info("=" * 60)

        # Lê a data limite da tela
        data_limite = await ler_data_limite_tela(page)

        # Usa o override de data (GUI) ou pergunta interativamente (CLI)
        if data_inicio_override:
            data_inicio_fase = data_inicio_override
            log.info(f"Data de início: {data_inicio_fase}")
        else:
            data_inicio_fase = perguntar_data_inicio(config.DATA_INICIO)

        # Gera os meses a partir da data escolhida
        todos_meses = gerar_meses(data_inicio_fase, data_limite)
        log.info(f"Período selecionado: {data_inicio_fase} até {data_limite} ({len(todos_meses)} meses)")

        ja_solicitados = set(progresso.get("solicitacoes_criadas", []))
        pendentes = [(i, d) for d in todos_meses if (i := f"{d[0]}_{d[1]}") not in ja_solicitados]

        log.info(f"Já solicitados anteriormente: {len(ja_solicitados)}")
        log.info(f"Pendentes nesta execução: {len(pendentes)}")

        # Registra no stats
        stats.data_inicio_fase1 = data_inicio_fase
        stats.data_limite = data_limite
        stats.f1_total_meses = len(todos_meses)
        stats.f1_ja_no_progresso = len(ja_solicitados)

        # Notifica GUI com totais
        if callback:
            callback({"tipo": "f1_inicio", "total": len(todos_meses), "pendentes": len(pendentes)})

        if not pendentes:
            log.info("✅ Todas as solicitações já foram criadas! Pule para a Fase 2.")
            if callback:
                callback({"tipo": "f1_todos_solicitados"})
            return

        pedidos_hoje = 0
        erros_consecutivos = 0
        ja_existia_count = 0

        for chave, (data_inicio, data_fim) in pendentes:
            # Limite diário (só conta pedidos novos, não os "já existe")
            if pedidos_hoje >= config.LIMITE_PEDIDOS_DIA:
                log.warning(f"⚠️ Limite de {config.LIMITE_PEDIDOS_DIA} pedidos/dia atingido. Encerre e retome amanhã.")
                stats.f1_limite_diario_atingido = True
                break

            # Tenta criar a solicitação (com retry apenas para "falha", não para "ja_existe")
            resultado = "falha"
            for tentativa in range(1, config.MAX_TENTATIVAS + 1):
                try:
                    resultado = await criar_solicitacao(page, data_inicio, data_fim, tentativa)
                    if resultado in ("sucesso", "ja_existe"):
                        if tentativa > 1:
                            stats.f1_tentativas_extras += (tentativa - 1)
                        erros_consecutivos = 0
                        break
                except RuntimeError as e:
                    err = str(e).lower()
                    if "browser_fechado" in err:
                        if playwright:
                            _, page = await aguardar_reconexao(playwright)
                            stats.f1_sessoes_expiradas += 1
                        else:
                            log.error("🔴 Chrome fechado e sem referência ao playwright para reconectar.")
                        resultado = "falha"
                        break
                    elif "expirada" in err:
                        log.error("🔴 Sessão expirada! Aguardando novo login...")
                        input("  Faça o login no eSocial e pressione ENTER para continuar...")
                        erros_consecutivos = 0
                        stats.f1_sessoes_expiradas += 1
                        resultado = "falha"
                        break
                if resultado == "falha" and tentativa < config.MAX_TENTATIVAS:
                    stats.f1_tentativas_extras += 1
                    log.info(f"  🔄 Aguardando {config.PAUSA_APOS_ERRO}s antes de nova tentativa...")
                    await asyncio.sleep(config.PAUSA_APOS_ERRO)

            if resultado == "sucesso":
                progresso["solicitacoes_criadas"].append(chave)
                salvar_progresso(progresso, cnpj)
                pedidos_hoje += 1
                stats.f1_novos_sucesso += 1
                log.info(f"  📊 Progresso: {pedidos_hoje} novos pedidos | {len(progresso['solicitacoes_criadas'])} total")
                if callback:
                    callback({"tipo": "f1_progresso", "atual": len(progresso["solicitacoes_criadas"]), "total": len(todos_meses)})

            elif resultado == "ja_existe":
                progresso["solicitacoes_criadas"].append(chave)
                salvar_progresso(progresso, cnpj)
                ja_existia_count += 1
                stats.f1_ja_existia += 1
                log.info(f"  📊 Pulados por já existirem: {ja_existia_count}")
                if callback:
                    callback({"tipo": "f1_progresso", "atual": len(progresso["solicitacoes_criadas"]), "total": len(todos_meses)})

            else:  # "falha"
                erros_consecutivos += 1
                stats.f1_falhas.append(chave)
                log.error(f"  ❌ Falha definitiva: {data_inicio} → {data_fim} (será tentado novamente na próxima execução)")
                if erros_consecutivos >= 5:
                    log.error("🔴 5 erros consecutivos. Interrompendo para diagnóstico.")
                    stats.f1_interrompido_por_erros = True
                    break

            # Pausa entre solicitações
            await asyncio.sleep(config.PAUSA_ENTRE_SOLICITACOES)

        log.info(f"\n✅ Fase 1 concluída.")
        log.info(f"   Novos pedidos criados : {pedidos_hoje}")
        log.info(f"   Já existiam (pulados) : {ja_existia_count}")
        log.info(f"   Total no progresso    : {len(progresso['solicitacoes_criadas'])}/{len(todos_meses)} meses")

    finally:
        if cb_handler:
            log.removeHandler(cb_handler)


# ─── FASE 2: Download dos XMLs ───────────────────────────────────────────────

async def verificar_downloads_disponiveis(page) -> int:
    """
    Verifica quantos arquivos estão disponíveis para download na tela de Consulta.
    Retorna o número de itens com link de download (class='icone-baixar').
    """
    try:
        await page.goto(
            "https://www.esocial.gov.br/portal/download/Pedido/Consulta",
            timeout=config.TIMEOUT_PAGINA,
            wait_until="domcontentloaded",
        )
        await page.wait_for_selector("#TipoConsulta", state="visible", timeout=10000)
        await page.select_option("#TipoConsulta", value="1")
        async with page.expect_navigation(wait_until="domcontentloaded", timeout=config.TIMEOUT_POS_SALVAR):
            await page.evaluate("document.querySelector('input[type=submit]').click()")
        await page.wait_for_selector("table.table-paginada tbody tr", timeout=15000)
        await page.wait_for_timeout(1500)
        # Expande DataTables com fallback
        await page.evaluate("""
            () => {
                try { jQuery('table.table-paginada').DataTable().page.len(-1).draw(false); return; } catch(e) {}
                try { jQuery.fn.dataTable.tables({ api: true }).page.len(-1).draw(false); return; } catch(e) {}
                try {
                    document.querySelectorAll('select[name$="_length"]').forEach(sel => {
                        if (!Array.from(sel.options).some(o => o.value==='-1')) {
                            const o = document.createElement('option');
                            o.value='-1'; o.text='Todos'; sel.appendChild(o);
                        }
                        sel.value='-1';
                        sel.dispatchEvent(new Event('change',{bubbles:true}));
                    });
                } catch(e) {}
            }
        """)
        await page.wait_for_timeout(2000)
        count = await page.evaluate("() => document.querySelectorAll('a.icone-baixar').length")
        return count
    except Exception:
        return 0


async def fase2_baixar_xmls(page, progresso: dict, stats: SessionStats,
                             callback=None, cnpj: str = None):
    """
    FASE 2: Acessa Downloads > Consulta, coleta todos os idPedido com
    situação 'Disponível para Baixar' e efetua o download de cada um.
    Se cnpj fornecido, salva em PASTA_DOWNLOAD/CNPJ/.
    """
    cb_handler = None
    if callback:
        cb_handler = CallbackLogHandler(callback)
        log.addHandler(cb_handler)

    try:
        log.info("=" * 60)
        log.info("  FASE 2 — BAIXANDO XMLs DISPONÍVEIS")
        log.info("=" * 60)

        # Pasta de destino: base ou base/CNPJ
        if cnpj:
            cnpj_limpo = re.sub(r'\D', '', str(cnpj))
            pasta_destino = Path(config.PASTA_DOWNLOAD) / cnpj_limpo
        else:
            pasta_destino = Path(config.PASTA_DOWNLOAD)
        pasta_destino.mkdir(parents=True, exist_ok=True)
        log.info(f"📁 Pasta de destino: {pasta_destino}")

        # ── Função auxiliar: extrai links do DOM atual ────────────────────────
        async def extrair_pedidos_disponiveis() -> list:
            return await page.evaluate("""
                () => {
                    const links = document.querySelectorAll('a.icone-baixar');
                    return Array.from(links).map(a => {
                        const href = a.getAttribute('href') || '';
                        const match = href.match(/idPedido=(\\d+)/);
                        const idPedido = match ? match[1] : null;
                        const tr = a.closest('tr');
                        const tds = tr ? Array.from(tr.querySelectorAll('td')) : [];
                        const detalhes = tds[3] ? tds[3].innerText.trim().replace(/\\s+/g, ' ') : '';
                        return { idPedido, detalhes };
                    }).filter(p => p.idPedido !== null);
                }
            """)

        async def consultar_e_expandir() -> list:
            """
            Navega para Downloads > Consulta, expande o DataTables e retorna
            todos os idPedido disponíveis que ainda não foram baixados.
            """
            await page.goto(
                "https://www.esocial.gov.br/portal/download/Pedido/Consulta",
                timeout=config.TIMEOUT_PAGINA,
                wait_until="domcontentloaded",
            )
            await page.wait_for_selector("#TipoConsulta", state="visible", timeout=15000)
            await page.select_option("#TipoConsulta", value="1")
            async with page.expect_navigation(
                wait_until="domcontentloaded",
                timeout=config.TIMEOUT_POS_SALVAR,
            ):
                await page.evaluate("document.querySelector('input[type=submit]').click()")

            await page.wait_for_selector("table.table-paginada tbody tr", timeout=15000)

            # Aguarda DataTables inicializar (polling até 10s)
            for _ in range(20):
                pronto = await page.evaluate("""
                    () => {
                        try { return jQuery.fn.dataTable.isDataTable('table.table-paginada'); }
                        catch(e) { return false; }
                    }
                """)
                if pronto:
                    break
                await page.wait_for_timeout(500)

            # Tenta expandir com 3 estratégias
            linhas_antes = await page.evaluate(
                "() => document.querySelectorAll('table.table-paginada tbody tr').length"
            )
            await page.evaluate("""
                () => {
                    try { jQuery('table.table-paginada').DataTable().page.len(-1).draw(false); return; } catch(e) {}
                    try { jQuery.fn.dataTable.tables({ api: true }).page.len(-1).draw(false); return; } catch(e) {}
                    try {
                        document.querySelectorAll('select[name$="_length"]').forEach(sel => {
                            if (!Array.from(sel.options).some(o => o.value==='-1')) {
                                const o = document.createElement('option');
                                o.value='-1'; o.text='Todos'; sel.appendChild(o);
                            }
                            sel.value='-1';
                            sel.dispatchEvent(new Event('change',{bubbles:true}));
                        });
                    } catch(e) {}
                }
            """)
            await page.wait_for_timeout(3000)

            linhas_depois = await page.evaluate(
                "() => document.querySelectorAll('table.table-paginada tbody tr').length"
            )
            if linhas_depois > linhas_antes:
                log.info(f"  📊 Tabela expandida: {linhas_antes} → {linhas_depois} linhas")
            else:
                log.warning(f"  ⚠️ Expansão incerta: {linhas_depois} linha(s) visível(is)")

            todos = await extrair_pedidos_disponiveis()
            ja = set(progresso.get("downloads_concluidos", []))
            pendentes = [p for p in todos if p["idPedido"] not in ja]
            log.info(f"  📋 Disponíveis: {len(todos)} | Já baixados: {len(ja)} | Pendentes: {len(pendentes)}")
            return pendentes

        # ── Loop principal: baixa até não restar nenhum pendente ─────────────
        novos_downloads = 0
        rodada = 0
        MAX_RODADAS = 10  # segurança contra loop infinito

        while rodada < MAX_RODADAS:
            rodada += 1
            log.info(f"  🔄 Rodada {rodada} — consultando pendentes...")
            pendentes_dl = await consultar_e_expandir()

            if not pendentes_dl:
                if rodada == 1:
                    log.info("  ✅ Nenhum arquivo pendente para download.")
                    stats.f2_ja_baixados = len(progresso.get("downloads_concluidos", []))
                else:
                    log.info(f"  ✅ Todos os arquivos baixados após {rodada - 1} rodada(s).")
                break

            # Atualiza total para callback de progresso
            if rodada == 1:
                stats.f2_disponiveis = len(pendentes_dl)
                if callback:
                    callback({"tipo": "f2_inicio", "total": len(pendentes_dl)})

            log.info(f"  ⬇️  Baixando {len(pendentes_dl)} arquivo(s) nesta rodada...")
            baixou_algum_nesta_rodada = False

            for pedido in pendentes_dl:
                id_pedido = pedido["idPedido"]
                detalhes  = pedido["detalhes"]

                periodo_nome = ""
                m = re.search(r"Inicial:\s*(\d{2})/(\d{2})/(\d{4})", detalhes)
                if m:
                    periodo_nome = f"{m.group(3)}-{m.group(2)}"

                log.info(f"  ⬇️  Baixando idPedido={id_pedido} | {detalhes[:50]}")
                url_download = f"https://www.esocial.gov.br/portal/Download/Pedido/Download?idPedido={id_pedido}"

                try:
                    async with page.expect_download(timeout=120000) as dl_info:
                        try:
                            await page.goto(url_download, timeout=config.TIMEOUT_PAGINA,
                                            wait_until="domcontentloaded")
                        except Exception as e:
                            if "download is starting" not in str(e).lower():
                                raise

                    download = await dl_info.value
                    nome_sugerido = download.suggested_filename or f"esocial_{periodo_nome}_{id_pedido}.zip"
                    if id_pedido not in nome_sugerido:
                        base, ext = os.path.splitext(nome_sugerido)
                        nome_sugerido = f"{base}_{id_pedido}{ext}"

                    destino = str(pasta_destino / nome_sugerido)
                    await download.save_as(destino)

                    log.info(f"  ✅ Salvo: {nome_sugerido}")
                    progresso["downloads_concluidos"].append(id_pedido)
                    salvar_progresso(progresso, cnpj)
                    novos_downloads += 1
                    stats.f2_baixados += 1
                    baixou_algum_nesta_rodada = True
                    if callback:
                        callback({"tipo": "f2_progresso", "baixados": novos_downloads,
                                  "total": stats.f2_disponiveis, "arquivo": nome_sugerido})
                    await asyncio.sleep(2)

                except Exception as e:
                    log.error(f"  ❌ Erro ao baixar idPedido={id_pedido}: {e}")
                    stats.f2_erros.append(f"idPedido={id_pedido} ({detalhes[:40]}): {e}")
                    continue

            # Se não baixou nada nesta rodada (todos falharam), evita loop infinito
            if not baixou_algum_nesta_rodada:
                log.error("  🔴 Nenhum arquivo baixado nesta rodada — encerrando para evitar loop.")
                break

        stats.f2_ja_baixados = len(set(progresso.get("downloads_concluidos", [])))
        log.info(f"\n✅ Fase 2 concluída.")
        log.info(f"   Baixados nesta sessão : {novos_downloads}")
        log.info(f"   Erros                 : {len(stats.f2_erros)}")
        log.info(f"   Total no progresso    : {len(progresso['downloads_concluidos'])}")

    finally:
        if cb_handler:
            log.removeHandler(cb_handler)


def gerar_relatorio(stats: SessionStats):
    """
    Gera um arquivo de relatório .txt com o resumo completo da execução.
    Salvo na mesma pasta do script com timestamp no nome.
    """
    stats.encerrar()
    nome_arquivo = f"relatorio_{stats.inicio.strftime('%Y%m%d_%H%M%S')}.txt"
    caminho = Path(__file__).parent / nome_arquivo

    linhas = []
    sep  = "=" * 62
    sep2 = "-" * 62

    def L(texto=""):
        linhas.append(texto)

    L(sep)
    L("  eSocial RPA — RELATÓRIO DE EXECUÇÃO")
    L(sep)
    L(f"  Gerado em  : {stats.fim.strftime('%d/%m/%Y %H:%M:%S')}")
    L(f"  Início     : {stats.inicio.strftime('%d/%m/%Y %H:%M:%S')}")
    L(f"  Término    : {stats.fim.strftime('%d/%m/%Y %H:%M:%S')}")
    L(f"  Duração    : {stats.duracao}")
    L(f"  Operação   : {stats.opcao_escolhida}")
    L()

    # ── FASE 1 ────────────────────────────────────────────────────────────────
    if stats.f1_total_meses > 0 or stats.data_inicio_fase1:
        L(sep2)
        L("  FASE 1 — SOLICITAÇÕES")
        L(sep2)
        L(f"  Período solicitado : {stats.data_inicio_fase1} → {stats.data_limite}")
        L(f"  Total de meses     : {stats.f1_total_meses}")
        L(f"  Já no progresso    : {stats.f1_ja_no_progresso}  (ignorados nesta sessão)")
        L(f"  Pendentes          : {stats.f1_total_meses - stats.f1_ja_no_progresso}")
        L()
        L("  Resultado desta sessão:")
        L(f"    ✅ Criados com sucesso   : {stats.f1_novos_sucesso}")
        L(f"    ⏭️  Já existiam (pulados) : {stats.f1_ja_existia}")
        L(f"    ❌ Falhas definitivas    : {len(stats.f1_falhas)}")
        L(f"    🔄 Tentativas extras     : {stats.f1_tentativas_extras}")
        L(f"    📊 Taxa de aproveitamento: {stats.taxa_sucesso_f1}")
        L()

        if stats.f1_sessoes_expiradas:
            L(f"  ⚠️  Sessões expiradas durante execução: {stats.f1_sessoes_expiradas}x")

        if stats.f1_limite_diario_atingido:
            L("  ⚠️  Limite de 100 pedidos/dia atingido — execução interrompida.")
            L("      Execute novamente amanhã para continuar.")

        if stats.f1_interrompido_por_erros:
            L("  🔴 Execução interrompida por 5 erros consecutivos.")
            L("      Verifique a conectividade e o estado da sessão.")

        if stats.f1_falhas:
            L()
            L("  Períodos que falharam (tentar novamente):")
            for p in stats.f1_falhas:
                periodo = p.replace("_", " → ")
                L(f"    • {periodo}")

    # ── FASE 2 ────────────────────────────────────────────────────────────────
    if stats.f2_disponiveis > 0 or stats.f2_baixados > 0 or stats.f2_erros:
        L()
        L(sep2)
        L("  FASE 2 — DOWNLOADS")
        L(sep2)
        L(f"  Disponíveis no portal : {stats.f2_disponiveis}")
        L(f"  Baixados nesta sessão : {stats.f2_baixados}")
        L(f"  Já baixados (pulados) : {stats.f2_ja_baixados}")
        L(f"  Erros de download     : {len(stats.f2_erros)}")
        L(f"  Pasta de destino      : {config.PASTA_DOWNLOAD}")

        if stats.f2_erros:
            L()
            L("  Arquivos com erro:")
            for e in stats.f2_erros:
                L(f"    • {e}")

    # ── SITUAÇÃO GERAL ────────────────────────────────────────────────────────
    L()
    L(sep2)
    L("  SITUAÇÃO GERAL DO PROGRESSO")
    L(sep2)
    progresso = carregar_progresso()
    criadas = progresso.get("solicitacoes_criadas", [])
    baixados = progresso.get("downloads_concluidos", [])
    L(f"  Solicitações no progresso.json : {len(criadas)}")
    if criadas:
        L(f"    Primeira : {criadas[0].replace('_', ' → ')}")
        L(f"    Última   : {criadas[-1].replace('_', ' → ')}")
    L(f"  Downloads no progresso.json    : {len(baixados)}")

    # ── SUGESTÕES ─────────────────────────────────────────────────────────────
    L()
    L(sep2)
    L("  OBSERVAÇÕES E SUGESTÕES DE MELHORIA")
    L(sep2)

    sugestoes = []

    if len(stats.f1_falhas) > 0:
        sugestoes.append(
            f"  • {len(stats.f1_falhas)} período(s) falharam. Execute a Fase 1 novamente\n"
            f"    a partir do primeiro mês com falha para reprocessá-los."
        )
    if stats.f1_tentativas_extras > 10:
        sugestoes.append(
            f"  • Foram necessárias {stats.f1_tentativas_extras} tentativas extras.\n"
            f"    Considere aumentar PAUSA_ENTRE_SOLICITACOES no config.py\n"
            f"    se o portal estiver respondendo lentamente."
        )
    if stats.f1_limite_diario_atingido:
        sugestoes.append(
            "  • O limite diário foi atingido. Retome amanhã escolhendo\n"
            "    a Fase 1 e informando o mês seguinte ao último processado."
        )
    if stats.f2_erros:
        sugestoes.append(
            f"  • {len(stats.f2_erros)} download(s) falharam. Execute a Fase 2\n"
            f"    novamente — o robô pulará os já baixados e tentará os restantes.\n"
            f"    Lembre: arquivos ficam disponíveis por apenas 7 dias!"
        )
    if stats.f2_disponiveis == 0 and stats.opcao_escolhida in ("Fase 2", "Fase 1 + Fase 2"):
        sugestoes.append(
            "  • Nenhum arquivo estava disponível na Fase 2.\n"
            "    O processamento pelo eSocial é assíncrono — aguarde algumas\n"
            "    horas e execute a Fase 2 novamente."
        )
    if not sugestoes:
        sugestoes.append("  • Nenhuma pendência identificada. Execução concluída sem problemas.")

    for s in sugestoes:
        L(s)

    L()
    L(sep)
    L(f"  Arquivo de log detalhado: {config.ARQUIVO_LOG}")
    L(f"  Progresso salvo em      : {config.ARQUIVO_PROGRESSO}")
    L(sep)

    # Salva o arquivo
    conteudo = "\n".join(linhas)
    with open(caminho, "w", encoding="utf-8") as f:
        f.write(conteudo)

    # Exibe no console também
    print("\n" + conteudo)
    log.info(f"📄 Relatório salvo em: {caminho}")
    return caminho


# ─── Menu Principal ──────────────────────────────────────────────────────────

def limpar_progresso_periodo():
    """
    Remove do progresso.json os meses de um intervalo informado pelo usuário,
    permitindo que sejam reprocessados na próxima execução da Fase 1.
    """
    progresso = carregar_progresso()
    criadas = progresso.get("solicitacoes_criadas", [])

    if not criadas:
        print("\n  ℹ️  Nenhuma solicitação no progresso.json ainda.")
        return

    print()
    print("  ╔══════════════════════════════════════════════════════╗")
    print("  ║     LIMPAR PROGRESSO — REPROCESSAR PERÍODO          ║")
    print("  ╠══════════════════════════════════════════════════════╣")
    print(f"  ║  Total no progresso  : {len(criadas):<30}║")
    print(f"  ║  Primeiro registrado : {criadas[0].replace('_', ' → '):<30}║")
    print(f"  ║  Último registrado   : {criadas[-1].replace('_', ' → '):<30}║")
    print("  ╠══════════════════════════════════════════════════════╣")
    print("  ║  Informe o intervalo a REMOVER do progresso.         ║")
    print("  ║  Esses meses serão retentados na próxima Fase 1.     ║")
    print("  ║  Formato: MM/AAAA  (ex: 01/2018)                    ║")
    print("  ╚══════════════════════════════════════════════════════╝")
    print()

    def pedir_mes(label: str) -> date | None:
        while True:
            entrada = input(f"  {label}: ").strip()
            if not entrada:
                print("  ❌ Campo obrigatório.")
                continue
            # Aceita MM/AAAA ou DD/MM/AAAA
            if re.fullmatch(r"\d{2}/\d{4}", entrada):
                entrada = f"01/{entrada}"
            try:
                return datetime.strptime(entrada, "%d/%m/%Y").date()
            except ValueError:
                print(f"  ❌ Data inválida: {entrada}. Use MM/AAAA (ex: 01/2018)")

    data_de = pedir_mes("Mês inicial a remover (ex: 01/2018)")
    if not data_de:
        return
    data_ate = pedir_mes("Mês final   a remover (ex: 12/2023) [ENTER = mesmo mês]")
    if not data_ate:
        data_ate = data_de

    if data_ate < data_de:
        data_de, data_ate = data_ate, data_de

    # Identifica quais chaves do progresso caem no intervalo
    def chave_no_intervalo(chave: str) -> bool:
        """Ex: chave = '01/01/2018_31/01/2018' → extrai o mês inicial."""
        try:
            parte_inicio = chave.split("_")[0]  # '01/01/2018'
            d = datetime.strptime(parte_inicio, "%d/%m/%Y").date()
            return data_de <= d <= date(data_ate.year, data_ate.month,
                                         ultimo_dia_mes(data_ate.year, data_ate.month))
        except Exception:
            return False

    remover = [c for c in criadas if chave_no_intervalo(c)]

    if not remover:
        print(f"\n  ℹ️  Nenhum mês encontrado no progresso entre "
              f"{data_de.strftime('%m/%Y')} e {data_ate.strftime('%m/%Y')}.")
        return

    print()
    print(f"  Os seguintes {len(remover)} meses serão removidos do progresso:")
    for c in remover:
        print(f"    • {c.replace('_', ' → ')}")
    print()

    confirmacao = input(f"  Confirma a remoção? (s/N): ").strip().lower()
    if confirmacao != "s":
        print("  ❌ Operação cancelada.")
        return

    # Remove do progresso e salva
    antes = len(criadas)
    progresso["solicitacoes_criadas"] = [c for c in criadas if c not in remover]
    salvar_progresso(progresso)
    depois = len(progresso["solicitacoes_criadas"])

    print()
    print(f"  ✅ {antes - depois} meses removidos do progresso.")
    print(f"  📋 Progresso atualizado: {depois} solicitações registradas.")
    print(f"  ▶️  Na próxima execução da Fase 1, esses meses serão reprocessados.")
    log.info(f"Progresso limpo: {antes - depois} meses removidos "
             f"({data_de.strftime('%m/%Y')} → {data_ate.strftime('%m/%Y')})")


def exibir_menu():
    print("\n" + "=" * 60)
    print("  eSocial RPA - Download Automático de XMLs")
    print("=" * 60)
    print("  1. FASE 1 - Criar solicitações mensais")
    print("  2. FASE 2 - Baixar XMLs disponíveis")
    print("  3. Executar Fase 1 + Fase 2 em sequência")
    print("  4. Ver progresso atual")
    print("  5. Reprocessar período (limpar progresso de um intervalo)")
    print("  0. Sair")
    print("=" * 60)
    return input("  Escolha: ").strip()


def ver_progresso():
    progresso = carregar_progresso()
    criadas = progresso.get("solicitacoes_criadas", [])
    baixados = progresso.get("downloads_concluidos", [])
    print(f"\n  📊 Solicitações criadas: {len(criadas)}")
    if criadas:
        print(f"     Primeira: {criadas[0]}")
        print(f"     Última:   {criadas[-1]}")
    print(f"  📥 Downloads concluídos: {len(baixados)}")


async def main():
    # Instrução de inicialização
    print("\n" + "=" * 60)
    print("  PRÉ-REQUISITO: Chrome deve estar aberto com debugger.")
    print("  Execute este comando antes:")
    print()
    print('  "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"')
    print('  --remote-debugging-port=9222')
    print('  --user-data-dir="C:\\chrome_esocial"')
    print()
    print("  Em seguida, faça login no eSocial com o certificado.")
    print("=" * 60)
    input("  Pressione ENTER quando estiver pronto...")

    opcao = exibir_menu()

    if opcao == "0":
        print("Saindo...")
        return

    if opcao == "4":
        ver_progresso()
        return

    if opcao == "5":
        limpar_progresso_periodo()
        return

    # Mapeia a opção para nome legível no relatório
    nomes_opcao = {
        "1": "Fase 1 — Criar solicitações",
        "2": "Fase 2 — Baixar XMLs",
        "3": "Fase 1 + Fase 2 (sequencial)",
    }

    stats = SessionStats()
    stats.opcao_escolhida = nomes_opcao.get(opcao, f"Opção {opcao}")

    progresso = carregar_progresso()

    try:
        async with async_playwright() as playwright:
            browser = await conectar_chrome(playwright)
            page = await obter_aba_esocial(browser)

            try:
                context = browser.contexts[0]
                await context.set_default_timeout(config.TIMEOUT_PAGINA)
            except:
                pass

            if opcao in ("1", "3"):
                await fase1_criar_solicitacoes(page, progresso, stats, playwright)

            if opcao == "3":
                print("\n  Aguardando 5 segundos antes de iniciar a Fase 2...")
                await asyncio.sleep(5)

            if opcao in ("2", "3"):
                await fase2_baixar_xmls(page, progresso, stats)

    except Exception as e:
        log.error(f"Erro fatal na execução: {e}")
    finally:
        # Gera o relatório independente de sucesso ou falha
        log.info("🏁 Robô finalizado. Gerando relatório...")
        gerar_relatorio(stats)


if __name__ == "__main__":
    asyncio.run(main())
