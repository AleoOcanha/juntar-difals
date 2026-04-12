import PyPDF2 as pyf
from pathlib import Path
import pandas as pd
import re
import os
import sys
import threading
from PyPDF2 import PdfMerger
import tkinter as tk
from tkinter import filedialog, ttk
from PIL import Image, ImageTk


def resource_path(relative):
    """Retorna o caminho correto tanto em desenvolvimento quanto no .exe (PyInstaller)."""
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, relative)

# ── Paleta de cores ──────────────────────────────────────────────────
TEAL        = "#004F4A"
TEAL_LIGHT  = "#006B64"
EMERALD     = "#2ECC71"
LIME        = "#D4E157"
LIME_DARK   = "#BCD044"
SLATE       = "#4A5568"
LOG_BG      = "#F7FAFC"
DIVIDER     = "#E2E8F0"
CARD_BG     = "#FFFFFF"
APP_BG      = "#EDF2F7"
TEXT_DARK   = "#1A202C"
MUTED       = "#A0AEC0"
ERROR_RED   = "#E53E3E"
WARNING_OR  = "#ED8936"


class JuntarDifalsApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Juntar Difals")
        self.root.geometry("580x660")
        self.root.configure(bg=APP_BG)
        self.root.resizable(False, False)
        self.root.withdraw()  # Oculta até o splash fechar

        self.pasta_pdfs = None
        self._pulse_id  = None
        self._pulse_on  = False

        # Carrega ícone uma vez
        icone_path  = resource_path("icone para programas.png")
        self.icon16 = None
        self.icon80 = None
        if os.path.exists(icone_path):
            self.icon16 = ImageTk.PhotoImage(Image.open(icone_path).resize((32, 32),  Image.LANCZOS))
            self.icon80 = ImageTk.PhotoImage(Image.open(icone_path).resize((80, 80),  Image.LANCZOS))
            self.root.iconphoto(True, self.icon16)

        self._build_ui()
        self._mostrar_splash()

    # ── Construção da interface ──────────────────────────────────────

    def _build_ui(self):
        self._build_header()
        tk.Frame(self.root, bg=DIVIDER, height=1).pack(fill="x")
        content = tk.Frame(self.root, bg=APP_BG)
        content.pack(fill="both", expand=True, padx=20, pady=15)
        self._build_card_selecao(content)
        tk.Frame(content, bg=DIVIDER, height=1).pack(fill="x", pady=6)
        self._build_card_status(content)
        tk.Frame(content, bg=DIVIDER, height=1).pack(fill="x", pady=6)
        self._build_log(content)
        tk.Frame(content, bg=DIVIDER, height=1).pack(fill="x", pady=(6, 8))
        self._build_botao(content)

    def _build_header(self):
        header = tk.Frame(self.root, bg=TEAL, height=68)
        header.pack(fill="x")
        header.pack_propagate(False)

        inner = tk.Frame(header, bg=TEAL)
        inner.pack(fill="both", expand=True, padx=18)

        if self.icon16:
            tk.Label(inner, image=self.icon16, bg=TEAL).pack(side="left", pady=18)

        txt = tk.Frame(inner, bg=TEAL)
        txt.pack(side="left", padx=(10, 0), pady=10)
        tk.Label(txt, text="Juntar Difals",
                 font=("Arial", 17, "bold"), fg="white", bg=TEAL).pack(anchor="w")
        tk.Label(txt, text="Processamento de Documentos DIFAL",
                 font=("Arial", 8), fg="#A8D5D1", bg=TEAL).pack(anchor="w")

        tk.Label(inner, text="Alma³ Soluções Tecnológicas",
                 font=("Arial", 8), fg="#A8D5D1", bg=TEAL).pack(side="right", pady=22)

    def _build_card_selecao(self, parent):
        # Wrapper com borda lateral
        self._border_selecao = tk.Frame(parent, bg=DIVIDER)
        self._border_selecao.pack(fill="x", pady=(0, 4))

        body = tk.Frame(self._border_selecao, bg=CARD_BG)
        body.pack(side="right", fill="both", expand=True, padx=(4, 0), pady=1)

        tk.Label(body, text="PASTA DE ENTRADA",
                 font=("Arial", 8, "bold"), fg=SLATE, bg=CARD_BG).pack(anchor="w", padx=12, pady=(10, 4))

        row = tk.Frame(body, bg=CARD_BG)
        row.pack(fill="x", padx=12, pady=(0, 10))

        self.lbl_pasta = tk.Label(row, text="Nenhuma pasta selecionada",
                                  font=("Arial", 10), fg=MUTED, bg=CARD_BG, anchor="w")
        self.lbl_pasta.pack(side="left", fill="x", expand=True)

        tk.Button(row, text="Selecionar Pasta",
                  font=("Arial", 9, "bold"), fg=TEXT_DARK, bg=LIME,
                  activebackground=LIME_DARK, relief="flat", padx=12, pady=4,
                  cursor="hand2", command=self._selecionar_pasta).pack(side="right")

    def _build_card_status(self, parent):
        self._border_status = tk.Frame(parent, bg=DIVIDER)
        self._border_status.pack(fill="x", pady=(0, 4))

        body = tk.Frame(self._border_status, bg=CARD_BG)
        body.pack(side="right", fill="both", expand=True, padx=(4, 0), pady=1)

        tk.Label(body, text="STATUS DO PROCESSO",
                 font=("Arial", 8, "bold"), fg=SLATE, bg=CARD_BG).pack(anchor="w", padx=12, pady=(10, 2))

        self.lbl_status = tk.Label(body, text="Aguardando início...",
                                   font=("Arial", 12), fg=SLATE, bg=CARD_BG, anchor="w")
        self.lbl_status.pack(anchor="w", padx=12, pady=(2, 8))

        # Configura estilos da barra de progresso
        s = ttk.Style()
        s.theme_use("default")
        s.configure("Teal.Horizontal.TProgressbar",
                    troughcolor=DIVIDER, background=TEAL, thickness=7)
        s.configure("Emerald.Horizontal.TProgressbar",
                    troughcolor=DIVIDER, background=EMERALD, thickness=7)

        self.progress = ttk.Progressbar(body, orient="horizontal",
                                        mode="determinate",
                                        style="Teal.Horizontal.TProgressbar")
        self.progress.pack(fill="x", padx=12)

        self.lbl_pag = tk.Label(body, text="", font=("Arial", 8), fg=MUTED, bg=CARD_BG)
        self.lbl_pag.pack(anchor="e", padx=12, pady=(2, 10))

    def _build_log(self, parent):
        tk.Label(parent, text="LOG DE EXECUÇÃO",
                 font=("Arial", 8, "bold"), fg=SLATE, bg=APP_BG).pack(anchor="w")

        frame = tk.Frame(parent, bg=LOG_BG, bd=1, relief="solid")
        frame.pack(fill="both", expand=True, pady=(4, 0))

        self.log = tk.Text(frame, bg=LOG_BG, fg=SLATE,
                           font=("Courier New", 9),
                           relief="flat", padx=10, pady=8,
                           state="disabled", wrap="word", height=10)

        sb = tk.Scrollbar(frame, command=self.log.yview, bg=DIVIDER)
        self.log.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.log.pack(fill="both", expand=True)

        self.log.tag_configure("info",    foreground=TEAL)
        self.log.tag_configure("success", foreground=EMERALD)
        self.log.tag_configure("error",   foreground=ERROR_RED)
        self.log.tag_configure("warning", foreground=WARNING_OR)
        self.log.tag_configure("default", foreground=SLATE)

    def _build_botao(self, parent):
        self.btn_iniciar = tk.Button(
            parent, text="▶   INICIAR PROCESSAMENTO",
            font=("Arial", 11, "bold"), fg=TEXT_DARK, bg=LIME,
            activebackground=LIME_DARK, relief="flat",
            padx=20, pady=10, cursor="hand2",
            state="disabled", command=self._iniciar_processamento
        )
        self.btn_iniciar.pack(fill="x")

    # ── Splash screen ────────────────────────────────────────────────

    def _mostrar_splash(self):
        splash = tk.Toplevel(self.root)
        splash.title("Juntar Difals")
        splash.geometry("380x280")
        splash.resizable(False, False)
        splash.configure(bg=CARD_BG)
        splash.grab_set()

        splash.update_idletasks()
        x = (splash.winfo_screenwidth()  // 2) - 190
        y = (splash.winfo_screenheight() // 2) - 140
        splash.geometry(f"380x280+{x}+{y}")

        if self.icon80:
            splash.iconphoto(True, self.icon16)
            tk.Label(splash, image=self.icon80, bg=CARD_BG).pack(pady=(20, 4))

        tk.Label(splash, text="Juntar Difals",
                 font=("Arial", 16, "bold"), fg=TEAL, bg=CARD_BG).pack()
        tk.Label(splash, text="Programa iniciado! Clique em OK para continuar.",
                 font=("Arial", 9), fg=SLATE, bg=CARD_BG).pack(pady=8)

        def _fechar():
            splash.destroy()
            self.root.deiconify()

        tk.Button(splash, text="OK", width=10,
                  font=("Arial", 10, "bold"), fg=TEXT_DARK, bg=LIME,
                  activebackground=LIME_DARK, relief="flat",
                  padx=10, pady=6, cursor="hand2",
                  command=_fechar).pack(pady=6)

        tk.Label(splash, text="Alma³ Soluções Tecnológicas",
                 font=("Arial", 8), fg=MUTED, bg=CARD_BG).place(
                 relx=1.0, rely=1.0, anchor="se", x=-8, y=-6)

    # ── Lógica da interface ──────────────────────────────────────────

    def _selecionar_pasta(self):
        pasta = filedialog.askdirectory(title="Selecione a pasta com as DIFALs")
        if pasta:
            self.pasta_pdfs = pasta
            exibir = pasta if len(pasta) <= 52 else "..." + pasta[-49:]
            self.lbl_pasta.config(text=exibir, fg=TEXT_DARK)
            self.btn_iniciar.config(state="normal")
            self._log(f"Pasta selecionada: {pasta}", "info")

    def _log(self, msg, tag="default"):
        self.log.config(state="normal")
        self.log.insert("end", msg + "\n", tag)
        self.log.see("end")
        self.log.config(state="disabled")

    def _set_status(self, texto, estado="idle"):
        cores_borda = {
            "idle":    DIVIDER,
            "running": TEAL,
            "success": EMERALD,
            "error":   ERROR_RED,
        }
        self.lbl_status.config(text=texto)
        self._border_status.config(bg=cores_borda.get(estado, DIVIDER))

        if estado == "running":
            self._iniciar_pulse()
        else:
            self._parar_pulse()

    def _iniciar_pulse(self):
        self._pulse_on = not self._pulse_on
        self.lbl_status.config(fg=TEAL if self._pulse_on else TEAL_LIGHT)
        self._pulse_id = self.root.after(550, self._iniciar_pulse)

    def _parar_pulse(self):
        if self._pulse_id:
            self.root.after_cancel(self._pulse_id)
            self._pulse_id = None
        self.lbl_status.config(fg=SLATE)

    # ── Processamento (thread separada) ─────────────────────────────

    def _iniciar_processamento(self):
        self.btn_iniciar.config(state="disabled")
        self._set_status("Em Execução...", "running")
        self._border_selecao.config(bg=TEAL)
        self.progress.config(mode="indeterminate",
                              style="Teal.Horizontal.TProgressbar")
        self.progress.start(15)
        self._log("─" * 52, "default")
        self._log("▶  Iniciando processamento...", "info")

        threading.Thread(target=self._processar, daemon=True).start()

    def _processar(self):
        def ui(fn, *args):
            self.root.after(0, fn, *args)

        try:
            os.makedirs(r"C:\Juntar difals", exist_ok=True)
            arquivo_saida = r"C:\Juntar difals\difals.pdf"

            # ── Mesclar PDFs ────────────────────────────────────────
            ui(self._log, "  Mesclando arquivos PDF...", "info")
            pdfs = [f for f in os.listdir(self.pasta_pdfs)
                    if f.lower().endswith(".pdf")]

            if not pdfs:
                ui(self._log, "  ERRO: Nenhum PDF encontrado na pasta.", "error")
                ui(self._encerrar_erro, "Nenhum PDF encontrado.")
                return

            ui(self._log, f"  {len(pdfs)} arquivo(s) encontrado(s).", "default")

            merger = PdfMerger()
            for f in pdfs:
                merger.append(os.path.join(self.pasta_pdfs, f))
            merger.write(arquivo_saida)
            merger.close()
            ui(self._log, "  PDFs mesclados com sucesso.", "success")

            # ── Ler e processar páginas ──────────────────────────────
            arquivo   = pyf.PdfReader(arquivo_saida)
            total_pag = len(arquivo.pages)
            ui(self._log, f"  Total de páginas: {total_pag}", "info")
            ui(self._ativar_barra_determinada, total_pag)

            dados_extraidos = []
            for i, pagina in enumerate(arquivo.pages):
                texto      = pagina.extract_text()
                linhas     = texto.split("\n")
                codigo     = None
                valor      = None

                for linha in linhas:
                    if linha.strip().startswith("85"):
                        codigo = "".join(linha.split())
                        break

                idx = next((j for j, l in enumerate(linhas)
                             if "Total a recolher" in l), -1)
                if idx != -1:
                    for l in linhas[idx + 1: idx + 6]:
                        m = re.search(r"(\d+,\d{2})", l)
                        if m:
                            try:
                                valor = float(m.group(1).replace(",", "."))
                                break
                            except ValueError:
                                continue

                if codigo and valor is not None:
                    dados_extraidos.append({
                        "Código de Barras": codigo,
                        "Valor": valor
                    })

                pct = int(((i + 1) / total_pag) * 100)
                ui(self._atualizar_progresso, i + 1, total_pag, pct)

            # ── Gerar Excel ─────────────────────────────────────────
            if not dados_extraidos:
                ui(self._log, "  AVISO: Nenhum dado válido encontrado.", "warning")
                ui(self._encerrar_erro, "Nenhum dado extraído.")
                return

            df         = pd.DataFrame(dados_extraidos).drop_duplicates().reset_index(drop=True)
            destino    = r"C:\Juntar difals"
            nome_excel = "Relatorio_Difals.xlsx"
            caminho    = os.path.join(destino, nome_excel)
            os.makedirs(destino, exist_ok=True)
            df.to_excel(caminho, index=False)

            ui(self._encerrar_sucesso, len(df), caminho)

        except Exception as e:
            ui(self._log, f"  ERRO inesperado: {e}", "error")
            ui(self._encerrar_erro, str(e))

    def _ativar_barra_determinada(self, total):
        self.progress.stop()
        self.progress.config(mode="determinate", maximum=total, value=0)

    def _atualizar_progresso(self, atual, total, pct):
        self.progress.config(value=atual)
        self.lbl_pag.config(text=f"Página {atual} de {total}  ({pct}%)")

    def _encerrar_sucesso(self, qtd, caminho):
        self._parar_pulse()
        self.progress.config(style="Emerald.Horizontal.TProgressbar",
                             value=self.progress["maximum"])
        self.lbl_status.config(text="Concluído com sucesso!", fg=EMERALD)
        self._border_status.config(bg=EMERALD)
        self._border_selecao.config(bg=EMERALD)
        self._log(f"  ✔  {qtd} registro(s) salvos.", "success")
        self._log(f"  Arquivo: {caminho}", "success")
        self._log("─" * 52, "default")
        self.btn_iniciar.config(state="normal", text="▶   PROCESSAR NOVAMENTE")

    def _encerrar_erro(self, msg):
        self._parar_pulse()
        self.lbl_status.config(text="Erro no processamento.", fg=ERROR_RED)
        self._border_status.config(bg=ERROR_RED)
        self.progress.stop()
        self.progress.config(mode="determinate", value=0)
        self._log(f"  Encerrado com erro: {msg}", "error")
        self._log("─" * 52, "default")
        self.btn_iniciar.config(state="normal")


if __name__ == "__main__":
    root = tk.Tk()
    JuntarDifalsApp(root)
    root.mainloop()
