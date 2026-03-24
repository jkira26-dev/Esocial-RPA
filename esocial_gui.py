"""
==============================================================
  eSocial RPA — Interface Gráfica v3
  7 melhorias: painel de resumo, expiração, notificação
  Windows, filtro de log, dashboard, auto-verificação e
  relatório consolidado.
==============================================================
"""

import asyncio
import json
import os
import queue
import re
import subprocess
import sys
import threading
from datetime import datetime, timedelta
from pathlib import Path
from tkinter import (
    BooleanVar, END, Frame, Label,
    messagebox, Scrollbar, StringVar, Text, Tk, ttk,
)

# ─── Constantes ──────────────────────────────────────────────────────────────

ARQUIVO_FILA    = Path(__file__).parent / "fila_empresas.json"
ARQUIVO_FILA_DL = Path(__file__).parent / "fila_downloads.json"
CHROME_PORT     = 9222
DIAS_VALIDADE   = 7          # eSocial disponibiliza por 7 dias
INTERVALO_AUTO  = 15         # minutos entre verificações automáticas

# ─── Cores da marca ──────────────────────────────────────────────────────────

COR_PRIMARY      = "#2C4A86"   # azul principal
COR_PRIMARY_DARK = "#1F3763"   # azul escuro (header)
COR_ACCENT       = "#F4B400"   # amarelo destaque
COR_ACCENT2      = "#1EA7D7"   # ciano secundário
COR_BG           = "#F2F2F2"   # fundo geral
COR_BG_CARD      = "#FFFFFF"   # fundo de cards
COR_VERDE        = "#27a050"
COR_AZUL         = COR_PRIMARY
COR_CINZA        = "#888780"
COR_ERRO         = "#a32d2d"
COR_AVISO        = "#ba7517"

STATUS_AGUARDANDO  = "Aguardando"
STATUS_VERIFICANDO = "Verificando..."
STATUS_FASE1       = "Fase 1 — Solicitações"
STATUS_FASE2       = "Fase 2 — Downloads"
STATUS_CONCLUIDO   = "Concluído"
STATUS_ERRO        = "Erro"
STATUS_PULADO      = "Downloads direto"

# ─── Utilitários ─────────────────────────────────────────────────────────────

def formatar_cnpj(cnpj: str) -> str:
    c = re.sub(r"\D", "", str(cnpj))
    return f"{c[:2]}.{c[2:5]}.{c[5:8]}/{c[8:12]}-{c[12:]}" if len(c) == 14 else cnpj

def validar_cnpj(cnpj: str) -> bool:
    return bool(re.fullmatch(r"\d{14}", re.sub(r"\D", "", str(cnpj))))

def validar_periodo(s: str) -> bool:
    return bool(re.fullmatch(r"\d{2}/\d{4}", s.strip()))

def dias_desde(dt_str: str) -> int | None:
    """Dias decorridos desde uma data 'DD/MM/YYYY HH:MM:SS' ou ISO."""
    for fmt in ("%d/%m/%Y %H:%M:%S", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M:%S"):
        try:
            return (datetime.now() - datetime.strptime(dt_str, fmt)).days
        except ValueError:
            pass
    return None

def expiracao_label(inserido_em: str) -> tuple[str, str]:
    """Retorna (label, cor) para coluna Expira em."""
    d = dias_desde(inserido_em)
    if d is None:
        return "?", COR_CINZA
    restam = DIAS_VALIDADE - d
    if restam <= 0:
        return "Expirado!", COR_ERRO
    if restam == 1:
        return "⚠ Amanhã", COR_AVISO
    if restam <= 2:
        return f"⚠ {restam}d", COR_AVISO
    return f"{restam}d", COR_VERDE

def carregar_fila() -> list:
    if ARQUIVO_FILA.exists():
        try:
            with open(ARQUIVO_FILA, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return []
    return []

def salvar_fila(fila: list):
    with open(ARQUIVO_FILA, "w", encoding="utf-8") as f:
        json.dump(fila, f, indent=2, ensure_ascii=False)

def carregar_fila_dl() -> list:
    if ARQUIVO_FILA_DL.exists():
        try:
            with open(ARQUIVO_FILA_DL, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return []
    return []

def salvar_fila_dl(fila: list):
    with open(ARQUIVO_FILA_DL, "w", encoding="utf-8") as f:
        json.dump(fila, f, indent=2, ensure_ascii=False)

def chrome_esta_aberto() -> bool:
    """Deprecado: agora o robô abre seu próprio browser."""
    return True

def notificar_windows(titulo: str, mensagem: str):
    """Notificação nativa do Windows. Tenta plyer → win10toast → silencioso."""
    try:
        from plyer import notification
        notification.notify(title=titulo, message=mensagem,
                            app_name="eSocial RPA", timeout=8)
        return
    except Exception:
        pass
    try:
        from win10toast import ToastNotifier
        ToastNotifier().show_toast(titulo, mensagem, duration=8, threaded=True)
        return
    except Exception:
        pass

def ler_todos_progressos() -> list[dict]:
    """Lê todos os progresso_*.json da pasta e retorna lista de dicts."""
    pasta = Path(__file__).parent
    resultado = []
    for arq in pasta.glob("progresso*.json"):
        try:
            with open(arq, "r", encoding="utf-8") as f:
                d = json.load(f)
                d["_arquivo"] = arq.name
                cnpj_match = re.search(r"progresso_(\d+)\.json", arq.name)
                d["_cnpj"] = cnpj_match.group(1) if cnpj_match else "global"
                resultado.append(d)
        except Exception:
            pass
    return resultado


# ─── Janela principal ────────────────────────────────────────────────────────

class App(Tk):
    def __init__(self):
        super().__init__()
        self.title("eSocial RPA — Download Automático de XML")
        self.geometry("1020x720")
        self.minsize(880, 600)
        self.configure(bg=COR_BG)

        # ── Ícone da aplicação ────────────────────────────────────────
        icon_path = Path(__file__).parent / "icon.ico"
        if icon_path.exists():
            try:
                self.iconbitmap(str(icon_path))
            except Exception:
                pass

        # ── Tema ttk com as cores da marca ────────────────────────────
        self._aplicar_tema()

        self.fila: list[dict]    = carregar_fila()
        self.fila_dl: list[dict] = carregar_fila_dl()
        self.rodando     = False
        self.rodando_dl  = False
        self.fila_msgs: queue.Queue = queue.Queue()

        # Melhoria #4 — filtros de log (armazena linhas em memória)
        self.log_lines:    list[tuple[str, str]] = []  # (texto, nivel)
        self.log_lines_dl: list[tuple[str, str]] = []
        self.filtro_log    = StringVar(value="todos")
        self.filtro_log_dl = StringVar(value="todos")

        # Melhoria #6 — auto-verificação
        self.auto_verif    = BooleanVar(value=False)
        self._auto_job_id  = None  # id do after() agendado

        self._build_ui()
        self._atualizar_grid()
        self._atualizar_grid_dl()
        self._atualizar_painel()
        self._poll_msgs()
        self.protocol("WM_DELETE_WINDOW", self._ao_fechar)

    def _aplicar_tema(self):
        """Aplica tema visual com as cores da marca eSocial IOB."""
        style = ttk.Style(self)

        # Base: clam dá mais controle que default/vista
        style.theme_use("clam")

        # ── Notebook (abas) ───────────────────────────────────────────
        style.configure("TNotebook",
                        background=COR_BG,
                        borderwidth=0)
        style.configure("TNotebook.Tab",
                        background="#dde3ee",
                        foreground=COR_PRIMARY_DARK,
                        padding=[12, 5],
                        font=("Segoe UI", 9))
        style.map("TNotebook.Tab",
                  background=[("selected", COR_BG_CARD),
                               ("active",   "#eef1f8")],
                  foreground=[("selected", COR_PRIMARY_DARK)],
                  expand=[("selected", [1, 1, 1, 0])])

        # ── Botões ────────────────────────────────────────────────────
        style.configure("TButton",
                        background=COR_PRIMARY,
                        foreground="white",
                        borderwidth=0,
                        focusthickness=0,
                        font=("Segoe UI", 9),
                        padding=[10, 5])
        style.map("TButton",
                  background=[("active",   COR_PRIMARY_DARK),
                               ("disabled", "#b0b8cc"),
                               ("pressed",  COR_PRIMARY_DARK)],
                  foreground=[("disabled", "#888")])

        # ── Radiobutton / Checkbutton ─────────────────────────────────
        for widget in ("TRadiobutton", "TCheckbutton"):
            style.configure(widget,
                            background=COR_BG_CARD,
                            foreground="#333",
                            font=("Segoe UI", 9))
            style.map(widget,
                      background=[("active", COR_BG_CARD)],
                      foreground=[("active", COR_PRIMARY)])

        # ── Entry (campos de texto) ───────────────────────────────────
        style.configure("TEntry",
                        fieldbackground="white",
                        bordercolor="#c5cfe0",
                        lightcolor="#c5cfe0",
                        darkcolor="#c5cfe0",
                        padding=[4, 3])
        style.map("TEntry",
                  bordercolor=[("focus", COR_PRIMARY),
                                ("active", COR_PRIMARY)])

        # ── Progressbar ───────────────────────────────────────────────
        style.configure("TProgressbar",
                        troughcolor="#dde3ee",
                        background=COR_PRIMARY,
                        borderwidth=0,
                        thickness=8)

        # ── Treeview (grades) ─────────────────────────────────────────
        style.configure("Treeview",
                        background="white",
                        fieldbackground="white",
                        foreground="#222",
                        rowheight=22,
                        font=("Segoe UI", 9),
                        borderwidth=0)
        style.configure("Treeview.Heading",
                        background=COR_PRIMARY,
                        foreground="white",
                        font=("Segoe UI", 9, "bold"),
                        relief="flat",
                        padding=[4, 6])
        style.map("Treeview",
                  background=[("selected", "#d6e4f7")],
                  foreground=[("selected", COR_PRIMARY_DARK)])
        style.map("Treeview.Heading",
                  background=[("active", COR_PRIMARY_DARK)])

        # ── Scrollbar ─────────────────────────────────────────────────
        style.configure("TScrollbar",
                        background="#dde3ee",
                        troughcolor=COR_BG,
                        borderwidth=0,
                        arrowsize=12)
        style.map("TScrollbar",
                  background=[("active", COR_PRIMARY)])

    # ── Construção ───────────────────────────────────────────────────────────

    def _ao_fechar(self):
        """Garante que jobs agendados sejam cancelados antes de fechar."""
        if self._auto_job_id:
            self.after_cancel(self._auto_job_id)
        self.destroy()

    def _build_ui(self):
        # Barra de título com cores da marca
        topo = Frame(self, bg=COR_PRIMARY_DARK, height=48)
        topo.pack(fill="x")
        topo.pack_propagate(False)

        # Logo "B" pequeno na barra
        logo_frame = Frame(topo, bg=COR_PRIMARY, width=48, height=48)
        logo_frame.pack(side="left")
        logo_frame.pack_propagate(False)
        Label(logo_frame, text="B", bg=COR_PRIMARY, fg="white",
              font=("Segoe UI", 18, "bold")).pack(expand=True)

        # Linha amarela de destaque
        accent_bar = Frame(topo, bg=COR_ACCENT, width=4, height=48)
        accent_bar.pack(side="left")
        accent_bar.pack_propagate(False)

        Label(topo, text="eSocial RPA", bg=COR_PRIMARY_DARK, fg="white",
              font=("Segoe UI", 13, "bold")).pack(side="left", padx=(12, 0), pady=10)
        Label(topo, text="Download Automático de XML", bg=COR_PRIMARY_DARK,
              fg=COR_ACCENT2, font=("Segoe UI", 10)).pack(side="left", padx=(8, 0), pady=10)

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=10, pady=(8, 0))

        self._aba_inclusao(nb)
        self._aba_downloads(nb)
        self._aba_painel(nb)
        self._aba_historico(nb)

        # Barra de status
        sb = Frame(self, bg="#dde3ee", height=26)
        sb.pack(fill="x", side="bottom")
        sb.pack_propagate(False)

        # Linha de acento superior da barra de status
        Frame(self, bg=COR_PRIMARY, height=2).pack(fill="x", side="bottom")

        self.lbl_status_bar = Label(sb, text="Pronto.", bg="#dde3ee",
                                    fg=COR_PRIMARY_DARK,
                                    font=("Segoe UI", 9), anchor="w")
        self.lbl_status_bar.pack(side="left", padx=10)

    # ── ABA INCLUSÃO E PROCESSAMENTO ─────────────────────────────────────────

    def _aba_inclusao(self, nb):
        frm = Frame(nb, bg=COR_BG)
        nb.add(frm, text="  Inclusão e Processamento  ")

        esq = Frame(frm, bg=COR_BG, width=560)
        esq.pack(side="left", fill="both", expand=True, padx=(10, 4), pady=10)
        esq.pack_propagate(False)

        # Formulário
        card = Frame(esq, bg=COR_BG_CARD, highlightbackground="#c5cfe0", highlightthickness=1)
        card.pack(fill="x", pady=(0, 4))
        Label(card, text="Nova empresa", bg=COR_BG_CARD,
              font=("Segoe UI", 9, "bold"), fg=COR_PRIMARY_DARK).pack(anchor="w", padx=12, pady=(10, 4))

        row_tipo = Frame(card, bg=COR_BG_CARD); row_tipo.pack(fill="x", padx=12, pady=2)
        Label(row_tipo, text="Tipo de acesso:", bg=COR_BG_CARD, font=("Segoe UI", 9),
              fg=COR_PRIMARY_DARK, width=14, anchor="w").pack(side="left")
        self.var_tipo = StringVar(value="proprio")
        ttk.Radiobutton(row_tipo, text="Próprio (A1)", variable=self.var_tipo,
                        value="proprio", command=self._on_tipo_change).pack(side="left", padx=4)
        ttk.Radiobutton(row_tipo, text="Procuração", variable=self.var_tipo,
                        value="procuracao", command=self._on_tipo_change).pack(side="left", padx=4)

        r2 = Frame(card, bg=COR_BG_CARD); r2.pack(fill="x", padx=12, pady=2)
        self.lbl_cnpj_tipo = Label(r2, text="CNPJ da empresa:", bg=COR_BG_CARD,
                                   font=("Segoe UI", 9), fg=COR_PRIMARY_DARK, width=14, anchor="w")
        self.lbl_cnpj_tipo.pack(side="left")
        self.entry_cnpj = ttk.Entry(r2, width=22, font=("Segoe UI", 9))
        self.entry_cnpj.pack(side="left", padx=(0, 6))
        self.entry_cnpj.bind("<FocusOut>", lambda e: self._atualizar_resumo())
        Label(r2, text="Nome/Descrição:", bg=COR_BG_CARD,
              font=("Segoe UI", 9), fg=COR_PRIMARY_DARK).pack(side="left")
        self.entry_nome = ttk.Entry(r2, width=20, font=("Segoe UI", 9))
        self.entry_nome.pack(side="left", padx=(0, 6))

        r3 = Frame(card, bg=COR_BG_CARD); r3.pack(fill="x", padx=12, pady=2)
        Label(r3, text="Período:", bg=COR_BG_CARD, font=("Segoe UI", 9),
              fg=COR_PRIMARY_DARK, width=14, anchor="w").pack(side="left")
        self.entry_inicio = ttk.Entry(r3, width=10, font=("Segoe UI", 9))
        self.entry_inicio.insert(0, "01/2018"); self.entry_inicio.pack(side="left")
        Label(r3, text=" até ", bg=COR_BG_CARD, font=("Segoe UI", 9), fg="#888").pack(side="left")
        self.entry_fim = ttk.Entry(r3, width=10, font=("Segoe UI", 9))
        self.entry_fim.insert(0, "03/2026"); self.entry_fim.pack(side="left")
        Label(r3, text="  (MM/AAAA)", bg=COR_BG_CARD,
              font=("Segoe UI", 8), fg="#aaa").pack(side="left")

        row_btn = Frame(card, bg=COR_BG_CARD); row_btn.pack(fill="x", padx=12, pady=(10, 10))
        ttk.Button(row_btn, text="+ Inserir na fila",
                   command=self._inserir_empresa).pack(side="left")
        ttk.Button(row_btn, text="Remover selecionado",
                   command=self._remover_empresa).pack(side="left", padx=6)
        ttk.Button(row_btn, text="Mover para Downloads",
                   command=self._mover_selecionado_para_dl).pack(side="left")
        ttk.Button(row_btn, text="Limpar fila",
                   command=self._limpar_fila).pack(side="left", padx=6)

        # ── Melhoria #1: painel de resumo ────────────────────────────────────
        self.card_resumo = Frame(esq, bg=COR_BG_CARD,
                                 highlightbackground="#c5cfe0", highlightthickness=1)
        self.card_resumo.pack(fill="x", pady=(0, 6))
        Label(self.card_resumo, text="Resumo da empresa", bg=COR_BG_CARD,
              font=("Segoe UI", 9, "bold"), fg=COR_PRIMARY_DARK).pack(anchor="w", padx=12, pady=(8, 4))
        self.lbl_resumo = Label(self.card_resumo,
                                text="Digite o CNPJ acima para ver o resumo do progresso local.",
                                bg=COR_BG_CARD, font=("Segoe UI", 9), fg=COR_CINZA, anchor="w")
        self.lbl_resumo.pack(fill="x", padx=12, pady=(0, 8))

        # Grade
        Label(esq, text="Fila de processamento", bg=COR_BG,
              font=("Segoe UI", 9, "bold"), fg=COR_PRIMARY_DARK).pack(anchor="w", pady=(2, 2))

        cols = ("nr", "empresa", "cnpj", "periodo", "fase", "status")
        self.tree = ttk.Treeview(esq, columns=cols, show="headings", height=8)
        self.tree.heading("nr",      text="#",       anchor="center")
        self.tree.heading("empresa", text="Empresa", anchor="w")
        self.tree.heading("cnpj",    text="CNPJ",    anchor="w")
        self.tree.heading("periodo", text="Período", anchor="center")
        self.tree.heading("fase",    text="Fase",    anchor="center")
        self.tree.heading("status",  text="Status",  anchor="center")
        self.tree.column("nr",      width=30,  stretch=False, anchor="center")
        self.tree.column("empresa", width=160, stretch=True)
        self.tree.column("cnpj",    width=130, stretch=False)
        self.tree.column("periodo", width=110, stretch=False, anchor="center")
        self.tree.column("fase",    width=70,  stretch=False, anchor="center")
        self.tree.column("status",  width=130, stretch=False, anchor="center")
        self.tree.tag_configure("concluido",   foreground=COR_VERDE)
        self.tree.tag_configure("erro",        foreground=COR_ERRO)
        self.tree.tag_configure("processando", foreground=COR_AZUL)
        self.tree.tag_configure("aguardando",  foreground=COR_CINZA)
        vsb = Scrollbar(esq, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="left", fill="y")

        # Painel direito
        dir_ = Frame(frm, bg=COR_BG, width=370)
        dir_.pack(side="right", fill="both", padx=(4, 10), pady=10)
        dir_.pack_propagate(False)

        cc = Frame(dir_, bg=COR_BG_CARD, highlightbackground="#c5cfe0", highlightthickness=1)
        cc.pack(fill="x", pady=(0, 8))
        Label(cc, text="Acesso ao eSocial", bg=COR_BG_CARD,
              font=("Segoe UI", 9, "bold"), fg=COR_PRIMARY_DARK).pack(anchor="w", padx=12, pady=(8, 4))
        fc = Frame(cc, bg=COR_BG_CARD); fc.pack(fill="x", padx=12, pady=(0, 8))
        Label(fc, text="O robô abrirá o portal automaticamente ao iniciar.", 
              bg=COR_BG_CARD, font=("Segoe UI", 8), fg=COR_CINZA, wraplength=300, justify="left").pack(side="left")

        Label(dir_, text="Progresso geral", bg=COR_BG,
              font=("Segoe UI", 9, "bold"), fg=COR_PRIMARY_DARK).pack(anchor="w", pady=(0, 2))
        self.pb_geral = ttk.Progressbar(dir_, mode="determinate", maximum=100)
        self.pb_geral.pack(fill="x", pady=(0, 2))
        self.lbl_pb = Label(dir_, text="—", bg=COR_BG,
                            font=("Segoe UI", 8), fg=COR_CINZA)
        self.lbl_pb.pack(anchor="w", pady=(0, 4))
        self.pb_empresa = ttk.Progressbar(dir_, mode="determinate", maximum=100)
        self.pb_empresa.pack(fill="x", pady=(0, 2))
        self.lbl_pb_emp = Label(dir_, text="—", bg=COR_BG,
                                font=("Segoe UI", 8), fg=COR_CINZA)
        self.lbl_pb_emp.pack(anchor="w", pady=(0, 6))

        fa = Frame(dir_, bg=COR_BG); fa.pack(fill="x", pady=(0, 8))
        self.btn_iniciar = ttk.Button(fa, text="▶  Iniciar processamento",
                                      command=self._iniciar)
        self.btn_iniciar.pack(side="left", fill="x", expand=True)
        self.btn_pausar = ttk.Button(fa, text="⏸  Pausar",
                                     command=self._pausar, state="disabled")
        self.btn_pausar.pack(side="left", padx=(6, 0))

        # Log + filtro (#4)
        self._build_log_panel(dir_, "log_principal")

    # ── ABA DOWNLOADS ─────────────────────────────────────────────────────────

    def _aba_downloads(self, nb):
        frm = Frame(nb, bg=COR_BG)
        nb.add(frm, text="  ⬇️  Downloads  ")

        esq = Frame(frm, bg=COR_BG, width=560)
        esq.pack(side="left", fill="both", expand=True, padx=(10, 4), pady=10)
        esq.pack_propagate(False)

        card = Frame(esq, bg=COR_BG_CARD, highlightbackground="#c5cfe0", highlightthickness=1)
        card.pack(fill="x", pady=(0, 6))
        Label(card, text="Nova empresa para download", bg=COR_BG_CARD,
              font=("Segoe UI", 9, "bold"), fg=COR_PRIMARY_DARK).pack(anchor="w", padx=12, pady=(10, 4))

        r1 = Frame(card, bg=COR_BG_CARD); r1.pack(fill="x", padx=12, pady=2)
        Label(r1, text="Tipo de acesso:", bg=COR_BG_CARD, font=("Segoe UI", 9),
              fg=COR_PRIMARY_DARK, width=14, anchor="w").pack(side="left")
        self.var_tipo_dl = StringVar(value="proprio")
        ttk.Radiobutton(r1, text="Próprio (A1)", variable=self.var_tipo_dl,
                        value="proprio", command=self._on_tipo_dl).pack(side="left", padx=4)
        ttk.Radiobutton(r1, text="Procuração", variable=self.var_tipo_dl,
                        value="procuracao", command=self._on_tipo_dl).pack(side="left", padx=4)

        r2 = Frame(card, bg=COR_BG_CARD); r2.pack(fill="x", padx=12, pady=2)
        self.lbl_cnpj_dl = Label(r2, text="CNPJ da empresa:", bg=COR_BG_CARD,
                                 font=("Segoe UI", 9), fg=COR_PRIMARY_DARK, width=14, anchor="w")
        self.lbl_cnpj_dl.pack(side="left")
        self.entry_cnpj_dl = ttk.Entry(r2, width=22, font=("Segoe UI", 9))
        self.entry_cnpj_dl.pack(side="left", padx=(0, 6))
        Label(r2, text="Nome/Descrição:", bg=COR_BG_CARD,
              font=("Segoe UI", 9), fg=COR_PRIMARY_DARK).pack(side="left")
        self.entry_nome_dl = ttk.Entry(r2, width=20, font=("Segoe UI", 9))
        self.entry_nome_dl.pack(side="left", padx=(0, 6))

        r3 = Frame(card, bg=COR_BG_CARD); r3.pack(fill="x", padx=12, pady=2)
        Label(r3, text="Período:", bg=COR_BG_CARD, font=("Segoe UI", 9),
              fg=COR_PRIMARY_DARK, width=14, anchor="w").pack(side="left")
        self.entry_ini_dl = ttk.Entry(r3, width=10, font=("Segoe UI", 9))
        self.entry_ini_dl.insert(0, "01/2018"); self.entry_ini_dl.pack(side="left")
        Label(r3, text=" até ", bg=COR_BG_CARD, font=("Segoe UI", 9), fg="#888").pack(side="left")
        self.entry_fim_dl = ttk.Entry(r3, width=10, font=("Segoe UI", 9))
        self.entry_fim_dl.insert(0, "03/2026"); self.entry_fim_dl.pack(side="left")
        Label(r3, text="  (MM/AAAA)", bg=COR_BG_CARD,
              font=("Segoe UI", 8), fg="#aaa").pack(side="left")

        rb = Frame(card, bg=COR_BG_CARD); rb.pack(fill="x", padx=12, pady=(4, 10))
        ttk.Button(rb, text="+ Inserir na fila",
                   command=self._inserir_dl).pack(side="left")
        ttk.Button(rb, text="Remover selecionado",
                   command=self._remover_dl).pack(side="left", padx=6)
        ttk.Button(rb, text="Limpar fila",
                   command=self._limpar_fila_dl).pack(side="left")

        Label(esq, text="Fila de downloads", bg=COR_BG,
              font=("Segoe UI", 9, "bold"), fg=COR_PRIMARY_DARK).pack(anchor="w", pady=(2, 2))

        # Melhoria #2 — coluna Expira em
        cols = ("nr", "empresa", "cnpj", "periodo", "expira", "status")
        self.tree_dl = ttk.Treeview(esq, columns=cols, show="headings", height=10)
        self.tree_dl.heading("nr",      text="#",         anchor="center")
        self.tree_dl.heading("empresa", text="Empresa",   anchor="w")
        self.tree_dl.heading("cnpj",    text="CNPJ",      anchor="w")
        self.tree_dl.heading("periodo", text="Período",   anchor="center")
        self.tree_dl.heading("expira",  text="Expira em", anchor="center")
        self.tree_dl.heading("status",  text="Status",    anchor="center")
        self.tree_dl.column("nr",      width=30,  stretch=False, anchor="center")
        self.tree_dl.column("empresa", width=160, stretch=True)
        self.tree_dl.column("cnpj",    width=130, stretch=False)
        self.tree_dl.column("periodo", width=110, stretch=False, anchor="center")
        self.tree_dl.column("expira",  width=75,  stretch=False, anchor="center")
        self.tree_dl.column("status",  width=130, stretch=False, anchor="center")
        self.tree_dl.tag_configure("concluido",   foreground=COR_VERDE)
        self.tree_dl.tag_configure("erro",        foreground=COR_ERRO)
        self.tree_dl.tag_configure("processando", foreground=COR_AZUL)
        self.tree_dl.tag_configure("aguardando",  foreground=COR_CINZA)
        self.tree_dl.tag_configure("expirando",   foreground=COR_AVISO)
        self.tree_dl.tag_configure("expirado",    foreground=COR_ERRO)
        vsb2 = Scrollbar(esq, orient="vertical", command=self.tree_dl.yview)
        self.tree_dl.configure(yscrollcommand=vsb2.set)
        self.tree_dl.pack(side="left", fill="both", expand=True)
        vsb2.pack(side="left", fill="y")

        dir_ = Frame(frm, bg=COR_BG, width=370)
        dir_.pack(side="right", fill="both", padx=(4, 10), pady=10)
        dir_.pack_propagate(False)

        cc2 = Frame(dir_, bg=COR_BG_CARD, highlightbackground="#c5cfe0", highlightthickness=1)
        cc2.pack(fill="x", pady=(0, 8))
        Label(cc2, text="Acesso ao eSocial", bg=COR_BG_CARD,
              font=("Segoe UI", 9, "bold"), fg=COR_PRIMARY_DARK).pack(anchor="w", padx=12, pady=(8, 4))
        fc2 = Frame(cc2, bg=COR_BG_CARD); fc2.pack(fill="x", padx=12, pady=(0, 8))
        Label(fc2, text="O robô abrirá o portal automaticamente ao iniciar.", 
              bg=COR_BG_CARD, font=("Segoe UI", 8), fg=COR_CINZA, wraplength=300, justify="left").pack(side="left")

        Label(dir_, text="Progresso", bg=COR_BG,
              font=("Segoe UI", 9, "bold"), fg=COR_PRIMARY_DARK).pack(anchor="w", pady=(0, 2))
        self.pb_geral_dl = ttk.Progressbar(dir_, mode="determinate", maximum=100)
        self.pb_geral_dl.pack(fill="x", pady=(0, 2))
        self.lbl_pb_dl = Label(dir_, text="—", bg=COR_BG,
                               font=("Segoe UI", 8), fg=COR_CINZA)
        self.lbl_pb_dl.pack(anchor="w", pady=(0, 4))
        self.pb_empresa_dl = ttk.Progressbar(dir_, mode="determinate", maximum=100)
        self.pb_empresa_dl.pack(fill="x", pady=(0, 2))
        self.lbl_pb_emp_dl = Label(dir_, text="—", bg=COR_BG,
                                   font=("Segoe UI", 8), fg=COR_CINZA)
        self.lbl_pb_emp_dl.pack(anchor="w", pady=(0, 6))

        fa2 = Frame(dir_, bg=COR_BG); fa2.pack(fill="x", pady=(0, 4))
        self.btn_iniciar_dl = ttk.Button(fa2, text="▶  Iniciar downloads",
                                         command=self._iniciar_dl)
        self.btn_iniciar_dl.pack(side="left", fill="x", expand=True)
        self.btn_pausar_dl = ttk.Button(fa2, text="⏸  Pausar",
                                        command=self._pausar_dl, state="disabled")
        self.btn_pausar_dl.pack(side="left", padx=(6, 0))

        # Melhoria #6 — auto-verificação
        av_frame = Frame(dir_, bg=COR_BG_CARD, highlightbackground="#c5cfe0", highlightthickness=1)
        av_frame.pack(fill="x", pady=(0, 8))
        Label(av_frame, text="Verificação automática", bg=COR_BG_CARD,
              font=("Segoe UI", 9, "bold"), fg=COR_PRIMARY_DARK).pack(anchor="w", padx=12, pady=(8, 4))
        av_inner = Frame(av_frame, bg=COR_BG_CARD); av_inner.pack(fill="x", padx=12, pady=(0, 8))
        ttk.Checkbutton(av_inner, text=f"Verificar a cada {INTERVALO_AUTO} min",
                        variable=self.auto_verif,
                        command=self._toggle_auto_verif).pack(side="left")
        self.lbl_prox_verif = Label(av_inner, text="", bg=COR_BG_CARD,
                                    font=("Segoe UI", 8), fg=COR_CINZA)
        self.lbl_prox_verif.pack(side="left", padx=(8, 0))

        # Log + filtro (#4)
        self._build_log_panel(dir_, "log_dl")

    # ── ABA PAINEL GERAL ─────────────────────────────────────────────────────

    def _aba_painel(self, nb):
        frm = Frame(nb, bg=COR_BG)
        nb.add(frm, text="  📊  Painel  ")

        # Cards de totais
        row_cards = Frame(frm, bg=COR_BG)
        row_cards.pack(fill="x", padx=12, pady=(14, 8))

        def metric_card(parent, label):
            c = Frame(parent, bg=COR_BG_CARD, highlightbackground="#c5cfe0", highlightthickness=1)
            c.pack(side="left", expand=True, fill="x", padx=4)
            Label(c, text=label, bg=COR_BG_CARD, font=("Segoe UI", 8),
                  fg=COR_CINZA).pack(pady=(10, 2))
            lbl = Label(c, text="—", bg=COR_BG_CARD,
                        font=("Segoe UI", 20, "bold"), fg="#222")
            lbl.pack(pady=(0, 10))
            return lbl

        self.lbl_p_empresas  = metric_card(row_cards, "Empresas com progresso")
        self.lbl_p_solicit   = metric_card(row_cards, "Total de solicitações")
        self.lbl_p_baixados  = metric_card(row_cards, "XMLs baixados")
        self.lbl_p_pendentes = metric_card(row_cards, "Ainda sem download")

        # Alertas de expiração
        Label(frm, text="Alertas de expiração", bg=COR_BG,
              font=("Segoe UI", 9, "bold"), fg=COR_PRIMARY_DARK).pack(anchor="w", padx=14, pady=(4, 2))
        self.frame_alertas = Frame(frm, bg=COR_BG)
        self.frame_alertas.pack(fill="x", padx=14)

        # Lista de empresas
        Label(frm, text="Situação por empresa", bg=COR_BG,
              font=("Segoe UI", 9, "bold"), fg=COR_PRIMARY_DARK).pack(anchor="w", padx=14, pady=(10, 2))

        cols = ("cnpj", "solicitacoes", "baixados", "pendentes", "arquivo")
        self.tree_painel = ttk.Treeview(frm, columns=cols, show="headings", height=10)
        self.tree_painel.heading("cnpj",        text="CNPJ",              anchor="w")
        self.tree_painel.heading("solicitacoes", text="Solicitações",      anchor="center")
        self.tree_painel.heading("baixados",     text="XMLs baixados",     anchor="center")
        self.tree_painel.heading("pendentes",    text="Pendentes download",anchor="center")
        self.tree_painel.heading("arquivo",      text="Arquivo progresso", anchor="w")
        self.tree_painel.column("cnpj",        width=170, stretch=True)
        self.tree_painel.column("solicitacoes", width=100, stretch=False, anchor="center")
        self.tree_painel.column("baixados",     width=100, stretch=False, anchor="center")
        self.tree_painel.column("pendentes",    width=130, stretch=False, anchor="center")
        self.tree_painel.column("arquivo",      width=180, stretch=False)
        row_tree = Frame(frm, bg=COR_BG); row_tree.pack(fill="both", expand=True, padx=14)
        vsb3 = Scrollbar(row_tree, orient="vertical", command=self.tree_painel.yview)
        self.tree_painel.configure(yscrollcommand=vsb3.set)
        self.tree_painel.pack(in_=row_tree, side="left", fill="both", expand=True)
        vsb3.pack(in_=row_tree, side="left", fill="y")

        ttk.Button(frm, text="🔄  Atualizar painel",
                   command=self._atualizar_painel).pack(pady=8)

    # ── ABA HISTÓRICO ─────────────────────────────────────────────────────────

    def _aba_historico(self, nb):
        frm = Frame(nb, bg=COR_BG)
        nb.add(frm, text="  Histórico  ")

        Label(frm, text="Relatórios e progresso por empresa",
              bg=COR_BG, font=("Segoe UI", 10), fg=COR_CINZA).pack(pady=(30, 10))

        ttk.Button(frm, text="Abrir pasta de relatórios",
                   command=self._abrir_pasta_relatorios).pack(pady=4)

        # Melhoria #7 — relatório consolidado
        ttk.Button(frm, text="📄  Gerar relatório consolidado (todas as empresas)",
                   command=self._gerar_relatorio_consolidado).pack(pady=4)

    # ── Log com filtro (#4) ──────────────────────────────────────────────────

    def _build_log_panel(self, parent, chave: str):
        """Constrói painel de log com botões de filtro. chave = 'log_principal' ou 'log_dl'."""
        Label(parent, text="Log em tempo real", bg=COR_BG,
              font=("Segoe UI", 9, "bold"), fg=COR_PRIMARY_DARK).pack(anchor="w", pady=(0, 2))

        # Botões de filtro
        filtro_var = self.filtro_log if chave == "log_principal" else self.filtro_log_dl
        fb = Frame(parent, bg=COR_BG); fb.pack(fill="x", pady=(0, 3))
        for nivel, label in [("todos", "Todos"), ("ok", "Downloads"),
                              ("warn", "Avisos"), ("error", "Erros")]:
            ttk.Radiobutton(fb, text=label, variable=filtro_var, value=nivel,
                            command=lambda c=chave: self._refiltrar_log(c)).pack(side="left", padx=2)

        frame_log = Frame(parent, bg="#1e1e1e")
        frame_log.pack(fill="both", expand=True)
        txt = Text(frame_log, bg="#1e1e1e", fg="#d4d4d4", font=("Consolas", 8),
                   state="disabled", wrap="word", relief="flat")
        vsb = Scrollbar(frame_log, command=txt.yview)
        txt.configure(yscrollcommand=vsb.set)
        txt.tag_configure("ok",    foreground="#4ec9b0")
        txt.tag_configure("info",  foreground="#9cdcfe")
        txt.tag_configure("warn",  foreground="#dcdcaa")
        txt.tag_configure("error", foreground="#f44747")
        txt.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        if chave == "log_principal":
            self.txt_log = txt
        else:
            self.txt_log_dl = txt

        row_lc = Frame(parent, bg=COR_BG); row_lc.pack(fill="x", pady=(4, 0))
        ttk.Button(row_lc, text="Limpar log",
                   command=lambda: self._limpar_log(chave)).pack(side="right")

    def _refiltrar_log(self, chave: str):
        """Re-renderiza o widget de log aplicando o filtro atual."""
        if chave == "log_principal":
            linhas = self.log_lines
            filtro = self.filtro_log.get()
            txt    = self.txt_log
        else:
            linhas = self.log_lines_dl
            filtro = self.filtro_log_dl.get()
            txt    = self.txt_log_dl

        txt.config(state="normal")
        txt.delete("1.0", END)
        for linha, nivel in linhas:
            if filtro == "todos" or nivel == filtro:
                txt.insert(END, linha + "\n", nivel)
        txt.see(END)
        txt.config(state="disabled")

    def _limpar_log(self, chave: str):
        if chave == "log_principal":
            self.log_lines.clear()
            txt = self.txt_log
        else:
            self.log_lines_dl.clear()
            txt = self.txt_log_dl
        txt.config(state="normal")
        txt.delete("1.0", END)
        txt.config(state="disabled")

    # ── Formulário aba inclusão ───────────────────────────────────────────────

    def _on_tipo_change(self):
        self.lbl_cnpj_tipo.config(
            text="CNPJ representado:" if self.var_tipo.get() == "procuracao"
            else "CNPJ da empresa:")

    def _atualizar_resumo(self):
        """Melhoria #1 — lê progresso local e atualiza card de resumo."""
        cnpj = re.sub(r"\D", "", self.entry_cnpj.get())
        if not validar_cnpj(cnpj):
            self.lbl_resumo.config(
                text="Digite o CNPJ acima para ver o resumo do progresso local.",
                fg=COR_CINZA)
            return

        # Tenta ler progresso específico do CNPJ
        arq = Path(__file__).parent / f"progresso_{cnpj}.json"
        if not arq.exists():
            arq = Path(__file__).parent / "progresso.json"

        if not arq.exists():
            self.lbl_resumo.config(
                text=f"Nenhum progresso encontrado para {formatar_cnpj(cnpj)} — empresa nova.",
                fg=COR_AVISO)
            return

        try:
            with open(arq, "r", encoding="utf-8") as f:
                prog = json.load(f)
            n_sol  = len(prog.get("solicitacoes_criadas", []))
            n_dl   = len(prog.get("downloads_concluidos", []))
            n_pend = n_sol - n_dl
            inicio = self.entry_inicio.get().strip()
            fim    = self.entry_fim.get().strip()
            txt = (f"Progresso encontrado para {formatar_cnpj(cnpj)}:  "
                   f"{n_sol} solicitações criadas  ·  "
                   f"{n_dl} XMLs já baixados  ·  "
                   f"{n_pend} pendentes de download"
                   + (f"  |  Período configurado: {inicio} → {fim}" if inicio and fim else ""))
            cor = COR_VERDE if n_pend == 0 and n_sol > 0 else COR_AZUL
            self.lbl_resumo.config(text=txt, fg=cor)
        except Exception as e:
            self.lbl_resumo.config(text=f"Erro ao ler progresso: {e}", fg=COR_ERRO)

    def _inserir_empresa(self):
        cnpj  = re.sub(r"\D", "", self.entry_cnpj.get())
        nome  = self.entry_nome.get().strip() or formatar_cnpj(cnpj)
        inicio = self.entry_inicio.get().strip()
        fim    = self.entry_fim.get().strip()
        fase   = "fase1"  # Aba 1 é exclusiva para Fase 1

        if not validar_cnpj(cnpj):
            messagebox.showwarning("CNPJ inválido", "Informe um CNPJ válido com 14 dígitos.")
            return
        if not validar_periodo(inicio) or not validar_periodo(fim):
            messagebox.showwarning("Período inválido", "Use o formato MM/AAAA.")
            return

        empresa = {"tipo": self.var_tipo.get(), "cnpj": cnpj, "nome": nome,
                   "inicio": inicio, "fim": fim, "fase": fase,
                   "status": STATUS_AGUARDANDO}
        self.fila.append(empresa)
        salvar_fila(self.fila)
        self._atualizar_grid()
        self._log(f"Adicionado: {nome} ({formatar_cnpj(cnpj)})", "info")
        self.entry_cnpj.delete(0, END)
        self.entry_nome.delete(0, END)
        self.lbl_resumo.config(
            text="Digite o CNPJ acima para ver o resumo do progresso local.",
            fg=COR_CINZA)

    def _remover_empresa(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Seleção", "Selecione uma empresa na grade primeiro.")
            return
        idx = int(self.tree.item(sel[0])["values"][0]) - 1
        nome = self.fila[idx]["nome"]
        if messagebox.askyesno("Confirmar", f"Remover '{nome}' da fila?"):
            self.fila.pop(idx)
            salvar_fila(self.fila)
            self._atualizar_grid()

    def _mover_selecionado_para_dl(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("Seleção", "Selecione uma empresa na grade primeiro.")
            return
        idx = int(self.tree.item(sel[0])["values"][0]) - 1
        empresa = self.fila[idx]
        nome = empresa["nome"]
        
        if messagebox.askyesno("Confirmar Transferência", f"Deseja mover '{nome}' para a fila de downloads?"):
            empresa_dl = {
                "tipo": empresa.get("tipo", "proprio"),
                "cnpj": empresa.get("cnpj", ""),
                "nome": nome,
                "inicio": empresa.get("inicio", ""),
                "fim": empresa.get("fim", ""),
                "fase": "fase2",
                "status": STATUS_AGUARDANDO,
                "inserido_em": datetime.now().strftime("%d/%m/%Y %H:%M:%S")
            }
            if not any(e.get("cnpj") == empresa_dl["cnpj"] for e in self.fila_dl):
                self.fila_dl.append(empresa_dl)
                salvar_fila_dl(self.fila_dl)
                self._atualizar_grid_dl()
            
            self.fila.pop(idx)
            salvar_fila(self.fila)
            self._atualizar_grid()
            self._log(f"  📥 Empresa '{nome}' movida manualmente para a aba de Downloads.", "ok")

    def _limpar_fila(self):
        if not self.fila:
            return
        if messagebox.askyesno("Limpar fila", "Remover todas as empresas da fila?"):
            self.fila.clear()
            salvar_fila(self.fila)
            self._atualizar_grid()

    # ── Formulário aba downloads ──────────────────────────────────────────────

    def _on_tipo_dl(self):
        self.lbl_cnpj_dl.config(
            text="CNPJ representado:" if self.var_tipo_dl.get() == "procuracao"
            else "CNPJ da empresa:")

    def _inserir_dl(self):
        cnpj  = re.sub(r"\D", "", self.entry_cnpj_dl.get())
        nome  = self.entry_nome_dl.get().strip() or formatar_cnpj(cnpj)
        inicio = self.entry_ini_dl.get().strip()
        fim    = self.entry_fim_dl.get().strip()

        if not validar_cnpj(cnpj):
            messagebox.showwarning("CNPJ inválido", "Informe um CNPJ válido com 14 dígitos.")
            return
        if not validar_periodo(inicio) or not validar_periodo(fim):
            messagebox.showwarning("Período inválido", "Use o formato MM/AAAA.")
            return

        empresa = {"tipo": self.var_tipo_dl.get(), "cnpj": cnpj, "nome": nome,
                   "inicio": inicio, "fim": fim, "fase": "fase2",
                   "status": STATUS_AGUARDANDO,
                   "inserido_em": datetime.now().strftime("%d/%m/%Y %H:%M:%S")}
        self.fila_dl.append(empresa)
        salvar_fila_dl(self.fila_dl)
        self._atualizar_grid_dl()
        self._log_dl(f"Adicionado: {nome} ({formatar_cnpj(cnpj)})", "info")
        self.entry_cnpj_dl.delete(0, END)
        self.entry_nome_dl.delete(0, END)

    def _remover_dl(self):
        sel = self.tree_dl.selection()
        if not sel:
            messagebox.showinfo("Seleção", "Selecione uma empresa na grade primeiro.")
            return
        idx = int(self.tree_dl.item(sel[0])["values"][0]) - 1
        nome = self.fila_dl[idx]["nome"]
        if messagebox.askyesno("Confirmar", f"Remover '{nome}' da fila?"):
            self.fila_dl.pop(idx)
            salvar_fila_dl(self.fila_dl)
            self._atualizar_grid_dl()

    def _limpar_fila_dl(self):
        if not self.fila_dl:
            return
        if messagebox.askyesno("Limpar fila", "Remover todas as empresas da fila de downloads?"):
            self.fila_dl.clear()
            salvar_fila_dl(self.fila_dl)
            self._atualizar_grid_dl()

    # ── Chrome ────────────────────────────────────────────────────────────────

    def _verificar_chrome(self):
        # Deprecado: agora o login é feito via robô
        self._status_bar("O robô abrirá o portal ao iniciar.")

    def _verificar_chrome_dl(self):
        # Deprecado
        pass

    def _abrir_chrome(self):
        bat = Path(__file__).parent / "2_ABRIR_CHROME.bat"
        if bat.exists():
            subprocess.Popen(["cmd", "/c", str(bat)], shell=True)
        else:
            for p in [r"C:\Program Files\Google\Chrome\Application\chrome.exe",
                      r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"]:
                if Path(p).exists():
                    subprocess.Popen([p, f"--remote-debugging-port={CHROME_PORT}",
                                      "--user-data-dir=C:\\chrome_esocial",
                                      "https://esocial.gov.br"])
                    break
        self._log("Chrome sendo aberto... Faça o login e clique em Verificar.", "info")

    # ── Auto-verificação (#6) ─────────────────────────────────────────────────

    def _toggle_auto_verif(self):
        if self.auto_verif.get():
            self._log_dl("Verificação automática ativada.", "info")
            self._agendar_auto_verif()
        else:
            if self._auto_job_id:
                self.after_cancel(self._auto_job_id)
                self._auto_job_id = None
            self.lbl_prox_verif.config(text="")
            self._log_dl("Verificação automática desativada.", "warn")

    def _agendar_auto_verif(self):
        if not self.auto_verif.get():
            return
        proxima = datetime.now() + timedelta(minutes=INTERVALO_AUTO)
        self.lbl_prox_verif.config(
            text=f"próxima: {proxima.strftime('%H:%M')}")
        ms = INTERVALO_AUTO * 60 * 1000
        self._auto_job_id = self.after(ms, self._executar_auto_verif)

    def _executar_auto_verif(self):
        if not self.auto_verif.get():
            return
        if False: # Ignora check antigo
            self._log_dl("Auto-verificação: Chrome não está aberto — pulando.", "warn")
            self._agendar_auto_verif()
            return
        if self.rodando_dl:
            self._log_dl("Auto-verificação: download em andamento — pulando.", "warn")
            self._agendar_auto_verif()
            return

        self._log_dl("🔄 Auto-verificação iniciada...", "info")
        threading.Thread(target=self._thread_auto_verif, daemon=True).start()

    def _thread_auto_verif(self):
        try:
            asyncio.run(self._async_auto_verif())
        except Exception as e:
            self._log_dl(f"Auto-verificação erro: {e}", "error")
        finally:
            self.fila_msgs.put({"tipo": "agendar_auto_verif"})

    async def _async_auto_verif(self):
        import esocial_rpa as rpa
        sys.path.insert(0, str(Path(__file__).parent))
        from playwright.async_api import async_playwright

        async with async_playwright() as pw:
            browser = await rpa.conectar_chrome(pw)
            page    = await rpa.obter_aba_esocial(browser)
            qtd = await rpa.verificar_downloads_disponiveis(page)
            self._log_dl(f"🔍 Auto-verificação: {qtd} arquivo(s) disponível(is).", "info")
            if qtd > 0:
                notificar_windows("eSocial RPA — Arquivos disponíveis",
                                  f"{qtd} arquivo(s) prontos para download.")
                self._log_dl("🔔 Notificação enviada.", "ok")

    # ── Processamento — aba inclusão ──────────────────────────────────────────

    def _iniciar(self):
        if self.rodando:
            return
        if not self.fila:
            messagebox.showinfo("Fila vazia", "Adicione pelo menos uma empresa antes de iniciar.")
            return

        # O robô agora abre seu próprio browser via esocial_rpa.py

        for emp in self.fila:
            if emp["status"] not in (STATUS_CONCLUIDO,):
                emp["status"] = STATUS_AGUARDANDO
        salvar_fila(self.fila)
        self._atualizar_grid()

        self.rodando = True
        self.btn_iniciar.config(state="disabled")
        self.btn_pausar.config(state="normal")
        self._log("=" * 50, "info")
        self._log(f"Iniciando processamento de {len(self.fila)} empresa(s)...", "info")
        threading.Thread(target=self._thread_rpa, daemon=True).start()

    def _pausar(self):
        self.rodando = False
        self._log("Pausa solicitada...", "warn")
        self.btn_pausar.config(state="disabled")

    def _thread_rpa(self):
        try:
            asyncio.run(self._loop_empresas())
        except Exception as e:
            self.fila_msgs.put({"tipo": "log", "level": "error",
                                "text": f"Erro fatal: {e}", "ts": ""})
        finally:
            self.fila_msgs.put({"tipo": "concluido"})

    async def _loop_empresas(self):
        import esocial_rpa as rpa
        sys.path.insert(0, str(Path(__file__).parent))
        from playwright.async_api import async_playwright

        total = len(self.fila)
        async with async_playwright() as playwright:
            browser, page = await rpa.iniciar_browser_rpa(playwright)
            self._log("⏳ Aguardando login manual...", "warn")
            if not await rpa.aguardar_login_usuario(page):
                self._log("❌ Login não realizado.", "error")
                return
            rpa.ocultar_janela_browser(page)

            for idx, empresa in enumerate(self.fila):
                if not self.rodando:
                    break
                if empresa["status"] == STATUS_CONCLUIDO:
                    continue

                nome = empresa["nome"]
                cnpj = empresa.get("cnpj", "")
                fase = empresa["fase"]

                empresa["status"] = STATUS_VERIFICANDO
                self._msg_grid()
                self._msg_pb_geral(idx, total, nome)
                self._log(f"\n{'='*46}", "info")
                self._log(f"  Empresa {idx+1}/{total}: {nome} ({formatar_cnpj(cnpj)})", "info")
                self._log(f"{'='*46}", "info")

                stats = rpa.SessionStats()
                stats.opcao_escolhida = fase

                def cb(msg, _idx=idx):
                    self.fila_msgs.put({**msg, "_emp_idx": _idx})

                try:
                    if empresa["tipo"] == "procuracao" and cnpj:
                        empresa["status"] = "Trocando perfil..."
                        self._msg_grid()
                        await self._trocar_perfil(page, cnpj)

                    progresso = rpa.carregar_progresso(cnpj)

                    empresa["status"] = STATUS_FASE1
                    self._msg_grid()
                    ja = set(progresso.get("solicitacoes_criadas", []))
                    data_inicio_str = f"01/{empresa['inicio']}"
                    data_fim_str    = f"01/{empresa['fim']}"
                    todos = rpa.gerar_meses(data_inicio_str, data_fim_str)
                    pendentes_f1 = [d for d in todos if f"{d[0]}_{d[1]}" not in ja]

                    if not pendentes_f1:
                        self._log("  ✅ Todas as solicitações já existem.", "ok")
                    else:
                        await rpa.fase1_criar_solicitacoes(
                            page, progresso, stats,
                            playwright=playwright, callback=cb,
                            cnpj=cnpj or None, data_inicio_override=data_inicio_str)

                    empresa["status"] = STATUS_CONCLUIDO
                    self._msg_grid()
                    rpa.gerar_relatorio(stats)
                    self._log(f"  ✅ Solicitações de {nome} concluídas!", "ok")
                    
                    self._log(f"  ➡️ Transferindo empresa para a fila de Downloads...", "info")
                    empresa_dl = {
                        "tipo": empresa.get("tipo", "proprio"),
                        "cnpj": empresa.get("cnpj", ""),
                        "nome": nome,
                        "inicio": empresa.get("inicio", ""),
                        "fim": empresa.get("fim", ""),
                        "fase": "fase2",
                        "status": STATUS_AGUARDANDO,
                        "inserido_em": datetime.now().strftime("%d/%m/%Y %H:%M:%S")
                    }
                    self.fila_msgs.put({"tipo": "mover_para_dl", "empresa_dl": empresa_dl})

                except Exception as e:
                    empresa["status"] = STATUS_ERRO
                    self._msg_grid()
                    self._log(f"  ❌ Erro em {nome}: {e}", "error")
                finally:
                    salvar_fila(self.fila)

    # ── Processamento — aba downloads ─────────────────────────────────────────

    def _iniciar_dl(self):
        if self.rodando_dl:
            return
        if not self.fila_dl:
            messagebox.showinfo("Fila vazia", "Adicione pelo menos uma empresa antes de iniciar.")
            return

        # O login será feito após o 'Iniciar'

        for emp in self.fila_dl:
            if emp["status"] not in (STATUS_CONCLUIDO,):
                emp["status"] = STATUS_AGUARDANDO
        salvar_fila_dl(self.fila_dl)
        self._atualizar_grid_dl()

        self.rodando_dl = True
        self.btn_iniciar_dl.config(state="disabled")
        self.btn_pausar_dl.config(state="normal")
        self._log_dl("=" * 50, "info")
        self._log_dl(f"Iniciando downloads — {len(self.fila_dl)} empresa(s)...", "info")
        threading.Thread(target=self._thread_rpa_dl, daemon=True).start()

    def _pausar_dl(self):
        self.rodando_dl = False
        self._log_dl("Pausa solicitada...", "warn")
        self.btn_pausar_dl.config(state="disabled")

    def _thread_rpa_dl(self):
        try:
            asyncio.run(self._loop_empresas_dl())
        except Exception as e:
            self.fila_msgs.put({"tipo": "log_dl", "level": "error",
                                "text": f"Erro fatal: {e}", "ts": ""})
        finally:
            self.fila_msgs.put({"tipo": "concluido_dl"})

    async def _loop_empresas_dl(self):
        import esocial_rpa as rpa
        sys.path.insert(0, str(Path(__file__).parent))
        from playwright.async_api import async_playwright

        total = len(self.fila_dl)
        n_baixados_total = 0

        async with async_playwright() as playwright:
            browser, page = await rpa.iniciar_browser_rpa(playwright)
            self._log_dl("⏳ Aguardando login manual...", "warn")
            if not await rpa.aguardar_login_usuario(page):
                self._log_dl("❌ Login não realizado.", "error")
                return
            rpa.ocultar_janela_browser(page)

            for idx, empresa in enumerate(self.fila_dl):
                if not self.rodando_dl:
                    break
                if empresa["status"] == STATUS_CONCLUIDO:
                    continue

                nome = empresa["nome"]
                cnpj = empresa.get("cnpj", "")

                empresa["status"] = STATUS_VERIFICANDO
                self._msg_grid_dl()
                self._msg_pb_geral_dl(idx, total, nome)
                self._log_dl(f"\n{'='*46}", "info")
                self._log_dl(f"  Empresa {idx+1}/{total}: {nome} ({formatar_cnpj(cnpj)})", "info")
                self._log_dl(f"{'='*46}", "info")

                stats = rpa.SessionStats()
                stats.opcao_escolhida = "fase2"

                def cb(msg, _idx=idx):
                    self.fila_msgs.put({**msg, "_emp_idx": _idx, "destino": "dl"})

                try:
                    if empresa["tipo"] == "procuracao" and cnpj:
                        empresa["status"] = "Trocando perfil..."
                        self._msg_grid_dl()
                        await self._trocar_perfil(page, cnpj)

                    progresso = rpa.carregar_progresso(cnpj)

                    empresa["status"] = "Verificando downloads..."
                    self._msg_grid_dl()
                    self._log_dl("  🔍 Verificando arquivos disponíveis...", "info")
                    qtd = await rpa.verificar_downloads_disponiveis(page)
                    self._log_dl(f"  📦 {qtd} arquivo(s) disponível(is).", "info")

                    if qtd == 0:
                        self._log_dl("  ℹ️ Nenhum arquivo disponível ainda.", "warn")
                    else:
                        empresa["status"] = STATUS_FASE2
                        self._msg_grid_dl()
                        await rpa.fase2_baixar_xmls(
                            page, progresso, stats, callback=cb, cnpj=cnpj)
                        n_baixados_total += stats.f2_baixados

                    empresa["status"] = STATUS_CONCLUIDO
                    self._msg_grid_dl()
                    rpa.gerar_relatorio(stats)
                    self._log_dl(f"  ✅ {nome} concluído!", "ok")

                except Exception as e:
                    empresa["status"] = STATUS_ERRO
                    self._msg_grid_dl()
                    self._log_dl(f"  ❌ Erro em {nome}: {e}", "error")
                finally:
                    salvar_fila_dl(self.fila_dl)

        # Melhoria #3 — notificação ao concluir
        concluidas = sum(1 for e in self.fila_dl if e["status"] == STATUS_CONCLUIDO)
        notificar_windows(
            "eSocial RPA — Downloads concluídos",
            f"{concluidas} empresa(s) processadas · {n_baixados_total} XMLs baixados")

    # ── Troca de perfil ───────────────────────────────────────────────────────

    async def _trocar_perfil(self, page, cnpj: str):
        try:
            await page.goto("https://www.esocial.gov.br/portal/Home/Inicial",
                            timeout=30000, wait_until="domcontentloaded")
            await page.wait_for_timeout(1000)
            conteudo = await page.inner_text("body")
            if re.sub(r"\D", "", cnpj) in re.sub(r"\D", "", conteudo):
                return
            try:
                await page.click(
                    "a:has-text('Trocar Perfil'), a:has-text('Trocar Perfil/Módulo')",
                    timeout=5000)
                await page.wait_for_timeout(1000)
            except Exception:
                pass
            try:
                await page.wait_for_selector("input[name='cnpj'], #cnpj",
                                             state="visible", timeout=8000)
                campo = page.locator("input[name='cnpj'], #cnpj").first
                await campo.fill(re.sub(r"\D", "", cnpj))
                await page.keyboard.press("Enter")
                await page.wait_for_timeout(2000)
            except Exception:
                pass
        except Exception as e:
            self._log(f"  ⚠️ Aviso ao trocar perfil: {e}", "warn")

    # ── Grades ────────────────────────────────────────────────────────────────

    def _atualizar_grid(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        for i, emp in enumerate(self.fila):
            cnpj_fmt   = formatar_cnpj(emp["cnpj"]) if emp.get("cnpj") else "Próprio"
            per        = f"{emp['inicio']} – {emp['fim']}"
            fase_label = {"ambas": "F1+F2", "fase1": "F1", "fase2": "F2"}.get(emp["fase"], "F1+F2")
            status     = emp.get("status", STATUS_AGUARDANDO)
            tag = ("concluido" if status == STATUS_CONCLUIDO
                   else "erro" if status == STATUS_ERRO
                   else "processando" if any(k in status for k in
                        (STATUS_FASE1, STATUS_FASE2, "Trocando", "Verif"))
                   else "aguardando")
            self.tree.insert("", END,
                             values=(i+1, emp["nome"], cnpj_fmt, per, fase_label, status),
                             tags=(tag,))

    def _atualizar_grid_dl(self):
        for item in self.tree_dl.get_children():
            self.tree_dl.delete(item)
        for i, emp in enumerate(self.fila_dl):
            cnpj_fmt = formatar_cnpj(emp["cnpj"]) if emp.get("cnpj") else "Próprio"
            per      = f"{emp['inicio']} – {emp['fim']}"
            status   = emp.get("status", STATUS_AGUARDANDO)

            # Melhoria #2 — expiração
            inserido = emp.get("inserido_em", "")
            exp_label, exp_cor = expiracao_label(inserido) if inserido else ("?", COR_CINZA)

            tag = ("concluido" if status == STATUS_CONCLUIDO
                   else "erro" if status == STATUS_ERRO
                   else "processando" if any(k in status for k in
                        (STATUS_FASE2, "Trocando", "Verif"))
                   else "expirado" if exp_label == "Expirado!"
                   else "expirando" if "⚠" in exp_label
                   else "aguardando")

            self.tree_dl.insert("", END,
                                values=(i+1, emp["nome"], cnpj_fmt, per, exp_label, status),
                                tags=(tag,))

    def _msg_grid(self):
        self.fila_msgs.put({"tipo": "atualizar_grid"})

    def _msg_grid_dl(self):
        self.fila_msgs.put({"tipo": "atualizar_grid_dl"})

    def _msg_pb_geral(self, idx, total, nome):
        pct = int((idx / max(total, 1)) * 100)
        self.fila_msgs.put({"tipo": "pb_geral", "pct": pct,
                            "label": f"Empresa {idx+1}/{total}: {nome}"})

    def _msg_pb_geral_dl(self, idx, total, nome):
        pct = int((idx / max(total, 1)) * 100)
        self.fila_msgs.put({"tipo": "pb_geral_dl", "pct": pct,
                            "label": f"Empresa {idx+1}/{total}: {nome}"})

    # ── Painel (#5) ──────────────────────────────────────────────────────────

    def _atualizar_painel(self):
        progressos = ler_todos_progressos()

        total_sol  = sum(len(p.get("solicitacoes_criadas", [])) for p in progressos)
        total_dl   = sum(len(p.get("downloads_concluidos", [])) for p in progressos)
        total_pend = total_sol - total_dl

        self.lbl_p_empresas.config(text=str(len(progressos)))
        self.lbl_p_solicit.config(text=str(total_sol))
        self.lbl_p_baixados.config(text=str(total_dl))
        self.lbl_p_pendentes.config(text=str(total_pend),
                                    fg=COR_AVISO if total_pend > 0 else COR_VERDE)

        # Alertas de expiração na fila_dl
        for w in self.frame_alertas.winfo_children():
            w.destroy()

        alertas = [(e["nome"], e.get("inserido_em", ""), e.get("cnpj", ""))
                   for e in self.fila_dl if e.get("inserido_em")]
        if not alertas:
            Label(self.frame_alertas, text="Nenhuma empresa na fila de downloads.",
                  bg=COR_BG, font=("Segoe UI", 9), fg=COR_CINZA).pack(anchor="w")
        else:
            for nome, inserido, cnpj in alertas:
                label, cor = expiracao_label(inserido)
                frm = Frame(self.frame_alertas, bg=COR_BG_CARD,
                            highlightbackground="#c5cfe0", highlightthickness=1)
                frm.pack(fill="x", pady=2)
                Label(frm, text=f"{nome} ({formatar_cnpj(cnpj)})",
                      bg=COR_BG_CARD, font=("Segoe UI", 9), fg="#222").pack(side="left", padx=8, pady=4)
                Label(frm, text=label, bg=COR_BG_CARD,
                      font=("Segoe UI", 9, "bold"), fg=cor).pack(side="right", padx=8)

        # Grade por empresa
        for item in self.tree_painel.get_children():
            self.tree_painel.delete(item)
        for p in progressos:
            cnpj     = formatar_cnpj(p["_cnpj"]) if p["_cnpj"] != "global" else "global"
            n_sol    = len(p.get("solicitacoes_criadas", []))
            n_dl     = len(p.get("downloads_concluidos", []))
            n_pend   = n_sol - n_dl
            self.tree_painel.insert("", END,
                                    values=(cnpj, n_sol, n_dl, n_pend, p["_arquivo"]))

    # ── Relatório consolidado (#7) ────────────────────────────────────────────

    def _gerar_relatorio_consolidado(self):
        pasta = Path(__file__).parent
        relatorios = sorted(
            [r for r in pasta.glob("relatorio_*.txt")
             if "consolidado" not in r.name],
            reverse=True)
        progressos = ler_todos_progressos()

        if not relatorios and not progressos:
            messagebox.showinfo("Sem dados",
                                "Nenhum relatório ou progresso encontrado na pasta.")
            return

        ts  = datetime.now().strftime("%Y%m%d_%H%M%S")
        out = pasta / f"relatorio_consolidado_{ts}.txt"
        sep = "=" * 62

        linhas = [sep,
                  "  eSocial RPA — RELATÓRIO CONSOLIDADO",
                  f"  Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}",
                  f"  Empresas com progresso: {len(progressos)}",
                  sep, ""]

        # Resumo geral
        total_sol = sum(len(p.get("solicitacoes_criadas", [])) for p in progressos)
        total_dl  = sum(len(p.get("downloads_concluidos", [])) for p in progressos)
        linhas += [
            "  RESUMO GERAL",
            "-" * 62,
            f"  Total de solicitações criadas : {total_sol}",
            f"  Total de XMLs baixados        : {total_dl}",
            f"  Pendentes de download         : {total_sol - total_dl}",
            "",
        ]

        # Por empresa (via progresso.json)
        linhas += ["  SITUAÇÃO POR EMPRESA", "-" * 62]
        for p in progressos:
            cnpj   = formatar_cnpj(p["_cnpj"]) if p["_cnpj"] != "global" else "global"
            n_sol  = len(p.get("solicitacoes_criadas", []))
            n_dl   = len(p.get("downloads_concluidos", []))
            n_pend = n_sol - n_dl
            linhas.append(f"  CNPJ: {cnpj}")
            linhas.append(f"    Solicitações: {n_sol}  |  Baixados: {n_dl}  |  Pendentes: {n_pend}")
            linhas.append("")

        # Conteúdo dos relatórios individuais
        if relatorios:
            linhas += [sep, "  RELATÓRIOS INDIVIDUAIS", sep, ""]
            for r in relatorios[:20]:  # máx 20
                linhas.append(f"--- {r.name} ---")
                try:
                    with open(r, "r", encoding="utf-8") as f:
                        linhas.append(f.read())
                except Exception as e:
                    linhas.append(f"  Erro ao ler: {e}")
                linhas.append("")

        with open(out, "w", encoding="utf-8") as f:
            f.write("\n".join(linhas))

        messagebox.showinfo("Relatório gerado",
                            f"Relatório consolidado salvo em:\n{out.name}")
        if os.name == "nt":
            os.startfile(str(out))

    # ── Poll de mensagens ─────────────────────────────────────────────────────

    def _poll_msgs(self):
        try:
            while True:
                msg  = self.fila_msgs.get_nowait()
                tipo = msg.get("tipo", "")

                if tipo == "log":
                    nivel = {"info": "info", "warning": "warn",
                             "error": "error"}.get(msg.get("level", ""), "ok")
                    ts  = msg.get("ts", "")
                    txt = f"[{ts}] {msg.get('text', '')}"
                    self.log_lines.append((txt, nivel))
                    if len(self.log_lines) > 2000:
                        self.log_lines = self.log_lines[-1500:]
                    if self.filtro_log.get() in ("todos", nivel):
                        self._append_to_widget(self.txt_log, txt, nivel)

                elif tipo == "log_dl":
                    nivel = {"info": "info", "warning": "warn",
                             "error": "error"}.get(msg.get("level", ""), "ok")
                    ts  = msg.get("ts", "")
                    txt = f"[{ts}] {msg.get('text', '')}"
                    self.log_lines_dl.append((txt, nivel))
                    if len(self.log_lines_dl) > 2000:
                        self.log_lines_dl = self.log_lines_dl[-1500:]
                    if self.filtro_log_dl.get() in ("todos", nivel):
                        self._append_to_widget(self.txt_log_dl, txt, nivel)

                elif tipo == "mover_para_dl":
                    emp = msg.get("empresa_dl")
                    if emp:
                        if not any(e.get("cnpj") == emp.get("cnpj") for e in self.fila_dl):
                            self.fila_dl.append(emp)
                            salvar_fila_dl(self.fila_dl)
                            self._atualizar_grid_dl()

                elif tipo == "atualizar_grid":
                    self._atualizar_grid()
                elif tipo == "atualizar_grid_dl":
                    self._atualizar_grid_dl()

                elif tipo == "pb_geral":
                    self.pb_geral["value"] = msg["pct"]
                    self.lbl_pb.config(text=msg["label"])
                elif tipo == "pb_geral_dl":
                    self.pb_geral_dl["value"] = msg["pct"]
                    self.lbl_pb_dl.config(text=msg["label"])

                elif tipo == "f1_progresso":
                    pct = int((msg["atual"] / max(msg["total"], 1)) * 100)
                    self.pb_empresa["value"] = pct
                    self.lbl_pb_emp.config(
                        text=f"Solicitações: {msg['atual']}/{msg['total']} ({pct}%)")

                elif tipo == "f1_inicio":
                    self.pb_empresa["value"] = 0
                    self.lbl_pb_emp.config(text=f"Fase 1: 0/{msg['total']} solicitações")

                elif tipo == "f1_todos_solicitados":
                    self.pb_empresa["value"] = 100
                    self.lbl_pb_emp.config(text="Fase 1: todas já solicitadas")

                elif tipo == "f2_inicio":
                    dest = msg.get("destino", "")
                    pb = self.pb_empresa_dl if dest == "dl" else self.pb_empresa
                    lb = self.lbl_pb_emp_dl if dest == "dl" else self.lbl_pb_emp
                    pb["value"] = 0
                    lb.config(text=f"Fase 2: 0/{msg['total']} downloads")

                elif tipo == "f2_progresso":
                    dest = msg.get("destino", "")
                    pb = self.pb_empresa_dl if dest == "dl" else self.pb_empresa
                    lb = self.lbl_pb_emp_dl if dest == "dl" else self.lbl_pb_emp
                    pct = int((msg["baixados"] / max(msg["total"], 1)) * 100)
                    pb["value"] = pct
                    lb.config(text=f"Fase 2: {msg['baixados']}/{msg['total']} ({msg.get('arquivo','')})")

                elif tipo == "concluido":
                    self.rodando = False
                    self.btn_iniciar.config(state="normal")
                    self.btn_pausar.config(state="disabled")
                    self.pb_geral["value"] = 100
                    self.lbl_pb.config(text="Processamento concluído.")
                    self._log("\n✅ Processamento finalizado.", "ok")
                    self._status_bar("Processamento concluído.")
                    self._atualizar_grid()
                    self._atualizar_painel()
                    # Melhoria #3 — notificação
                    concluidas = sum(1 for e in self.fila if e["status"] == STATUS_CONCLUIDO)
                    notificar_windows("eSocial RPA — Concluído",
                                      f"{concluidas} empresa(s) processadas.")

                elif tipo == "concluido_dl":
                    self.rodando_dl = False
                    self.btn_iniciar_dl.config(state="normal")
                    self.btn_pausar_dl.config(state="disabled")
                    self.pb_geral_dl["value"] = 100
                    self.lbl_pb_dl.config(text="Downloads concluídos.")
                    self._log_dl("\n✅ Downloads finalizados.", "ok")
                    self._status_bar("Downloads concluídos.")
                    self._atualizar_grid_dl()
                    self._atualizar_painel()

                elif tipo == "agendar_auto_verif":
                    self._agendar_auto_verif()

        except queue.Empty:
            pass
        finally:
            self.after(150, self._poll_msgs)

    # ── Log ──────────────────────────────────────────────────────────────────

    def _log(self, texto: str, nivel: str = "info"):
        """Thread-safe: sempre envia via fila_msgs para ser renderizado na thread principal."""
        ts = datetime.now().strftime("%H:%M:%S")
        self.fila_msgs.put({"tipo": "log", "level": nivel,
                            "text": texto, "ts": ts})

    def _log_dl(self, texto: str, nivel: str = "info"):
        """Thread-safe: sempre envia via fila_msgs para ser renderizado na thread principal."""
        ts = datetime.now().strftime("%H:%M:%S")
        self.fila_msgs.put({"tipo": "log_dl", "level": nivel,
                            "text": texto, "ts": ts})

    def _append_to_widget(self, widget: Text, linha: str, nivel: str):
        tag = {"ok": "ok", "info": "info", "warn": "warn",
               "error": "error"}.get(nivel, "info")
        widget.config(state="normal")
        widget.insert(END, linha + "\n", tag)
        widget.see(END)
        widget.config(state="disabled")

    def _status_bar(self, texto: str):
        self.lbl_status_bar.config(text=texto)

    def _abrir_pasta_relatorios(self):
        pasta = str(Path(__file__).parent)
        if sys.platform == "win32":
            os.startfile(pasta)


# ─── Entrypoint ──────────────────────────────────────────────────────────────

def main():
    try:
        import playwright  # noqa: F401
    except ImportError:
        print("Playwright não instalado. Execute: 1_INSTALAR.bat")
        sys.exit(1)

    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
