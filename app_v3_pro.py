"""
Projeto: APA 7 Normalizer & Word Compiler - Desktop Edition (v3.0 Pro)
Autor: Aristófanes Paulo Gonfulo
Tecnologias: CustomTkinter + TkinterDnD + Win32 COM
Design: Modern Vibrant UI com Parser Heurístico de Folha de Rosto
"""

import os
import sys
import threading
import time
import re
import unicodedata
import customtkinter as ctk
from tkinter import filedialog, messagebox

# 1. Blindagem de streams para modo sem consola (PyInstaller --windowed)
if sys.stdout is None:
    sys.stdout = open(os.devnull, "w", encoding="utf-8")
if sys.stderr is None:
    sys.stderr = open(os.devnull, "w", encoding="utf-8")

# 2. Suporte nativo a Arrastar e Soltar (Drag & Drop)
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    HAS_DND = True
except ImportError:
    HAS_DND = False

import docx

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

# ---------------------------------------------------------------------------
# PARSER HEURÍSTICO INTELIGENTE DE METADADOS ACADÉMICOS
# ---------------------------------------------------------------------------
def extract_academic_metadata(docx_path):
    """Analisa semanticamente as linhas iniciais da folha de rosto."""
    try:
        doc = docx.Document(docx_path)
        clean_lines = []
        for p in doc.paragraphs:
            txt = unicodedata.normalize('NFC', p.text.replace('\xa0', ' ')).strip()
            if txt:
                clean_lines.append(txt)
            if len(clean_lines) >= 12:
                break

        if not clean_lines:
            return None

        first_title = clean_lines[0]
        cover_lines = []
        for idx, line in enumerate(clean_lines):
            # Interrompe se o título se repetir (início da página 3) ou ao atingir o Sumário
            if idx > 0 and (line == first_title or line.lower() in ["sumário", "sumario", "resumo", "abstract"]):
                break
            cover_lines.append(line)

        metadata = {
            "titulo": first_title,
            "autores": "",
            "instituicao": "",
            "curso": "",
            "disciplina": "",
            "docente": "",
            "data": ""
        }

        remaining = []
        for line in cover_lines[1:]:
            # Deteção de Data
            if re.search(r'\b\d{1,2}\s+de\s+[a-zçãé]+\s+de\s+\d{4}\b', line, re.I) or re.search(r'\b(19|20)\d{2}\b', line):
                metadata["data"] = line
            # Deteção de Docente / Professor / Orientador
            elif re.search(r'^(prof|dr|dra|msc|phd|orientador|docente|tutor|eng\.)\b', line, re.I):
                metadata["docente"] = line
            # Deteção de Instituição e Curso
            elif re.search(r'(instituto|universidade|faculdade|politécnico|politecnico|ispetec|isp\b|academia|escola superior)', line, re.I):
                if ',' in line:
                    parts = [pt.strip() for pt in line.split(',')]
                    for pt in parts:
                        if re.search(r'(instituto|universidade|faculdade|politécnico|politecnico|ispetec|isp\b)', pt, re.I):
                            metadata["instituicao"] = pt
                        else:
                            metadata["curso"] = pt
                else:
                    metadata["instituicao"] = line
            else:
                remaining.append(line)

        # Separação entre Autores e Disciplina / Cadeira
        tech_keywords = ['teste', 'qualidade', 'engenharia', 'algoritmo', 'programação', 'redes', 'sistemas', 'software', 'banco de dados', 'cálculo', 'disciplina', 'cadeira', 'informática']
        autores = []
        disciplinas = []

        for item in remaining:
            if any(k in item.lower() for k in tech_keywords) and not re.search(r'^[A-ZÀ-Ú][a-zà-ú]+(\s+[A-ZÀ-Ú][a-zà-ú]+){1,3}$', item):
                disciplinas.append(item)
            elif any(k in item.lower() for k in ['teste e qualidade', 'engenharia de software', 'sistemas de informação']):
                disciplinas.append(item)
            else:
                autores.append(item)

        metadata["autores"] = " & ".join(autores) if autores else "Não identificado"
        metadata["disciplina"] = " / ".join(disciplinas) if disciplinas else "Não identificado"
        if not metadata["instituicao"]:
            metadata["instituicao"] = "Não identificado"
        if not metadata["curso"]:
            metadata["curso"] = "Engenharia Informática"
        if not metadata["docente"]:
            metadata["docente"] = "Não identificado"
        if not metadata["data"]:
            metadata["data"] = "Não identificado"

        return metadata
    except Exception:
        return None

# ---------------------------------------------------------------------------
# INTERFACE GRÁFICA MODERNA
# ---------------------------------------------------------------------------
class ApaDesktopApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Formatador APA 7 & Compilador PDF (v3.0 Pro)")
        self.geometry("860x820")
        self.minsize(780, 720)

        if HAS_DND:
            try:
                TkinterDnD.require(self)
            except Exception:
                pass

        self.selected_file_path = None
        self.output_pdf_path = None

        self._build_ui()

    def _build_ui(self):
        # Contentor com barra de deslocamento para ecrãs menores
        self.scroll_container = ctk.CTkScrollableFrame(self, fg_color="transparent")
        self.scroll_container.pack(fill="both", expand=True, padx=15, pady=10)

        # 1. Cartão de Cabeçalho (Hero Header)
        self.header_card = ctk.CTkFrame(
            self.scroll_container,
            corner_radius=16,
            fg_color=("#EEF2FF", "#1E1B4B"),
            border_width=1,
            border_color=("#C7D2FE", "#312E81")
        )
        self.header_card.pack(fill="x", pady=(5, 12))

        # Tag de versão moderna
        self.badge_version = ctk.CTkLabel(
            self.header_card,
            text="✨ v3.0 Pro • Conformidade Estrita APA 7",
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color=("#4F46E5", "#A5B4FC"),
            fg_color=("#E0E7FF", "#312E81"),
            corner_radius=8,
            padx=10,
            pady=3
        )
        self.badge_version.pack(anchor="w", padx=20, pady=(15, 6))

        self.lbl_title = ctk.CTkLabel(
            self.header_card,
            text="Formatador & Compilador Académico",
            font=ctk.CTkFont(size=22, weight="bold"),
            text_color=("#1E293B", "#F8FAFC")
        )
        self.lbl_title.pack(anchor="w", padx=20, pady=(0, 2))

        self.lbl_subtitle = ctk.CTkLabel(
            self.header_card,
            text="Normalização programática OpenXML (ECMA-376) e renderização vetorial nativa em PDF.",
            font=ctk.CTkFont(size=12),
            text_color=("#64748B", "#94A3B8")
        )
        self.lbl_subtitle.pack(anchor="w", padx=20, pady=(0, 15))

        # 2. Zona de Seleção / Drop Zone
        self.drop_card = ctk.CTkFrame(
            self.scroll_container,
            corner_radius=16,
            fg_color=("#F8FAFC", "#0F172A"),
            border_width=2,
            border_color=("#6366F1", "#818CF8")
        )
        self.drop_card.pack(fill="x", pady=6)

        self.lbl_icon = ctk.CTkLabel(
            self.drop_card,
            text="📁",
            font=ctk.CTkFont(size=32)
        )
        self.lbl_icon.pack(pady=(12, 2))

        drop_msg = "Arraste e solte o seu ficheiro .docx aqui ou clique para selecionar" if HAS_DND else "Clique no botão abaixo para carregar o documento Word"

        self.lbl_file_status = ctk.CTkLabel(
            self.drop_card,
            text=drop_msg,
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=("#334155", "#E2E8F0")
        )
        self.lbl_file_status.pack(pady=2)

        self.btn_select_file = ctk.CTkButton(
            self.drop_card,
            text="Procurar Documento Word",
            command=self._select_file_dialog,
            font=ctk.CTkFont(size=12, weight="bold"),
            fg_color=("#4F46E5", "#6366F1"),
            hover_color=("#4338CA", "#4F46E5"),
            corner_radius=10,
            height=34
        )
        self.btn_select_file.pack(pady=(6, 14))

        if HAS_DND:
            try:
                self.drop_card.drop_target_register(DND_FILES)
                self.drop_card.dnd_bind("<<Drop>>", self._on_file_drop)
            except Exception:
                pass

        # 3. Painel de Pré-visualização com Badges Coloridos
        self.meta_card = ctk.CTkFrame(
            self.scroll_container,
            corner_radius=16,
            fg_color=("#FFFFFF", "#1E293B"),
            border_width=1,
            border_color=("#E2E8F0", "#334155")
        )
        self.meta_card.pack(fill="x", pady=8)

        self.lbl_meta_heading = ctk.CTkLabel(
            self.meta_card,
            text="Folha de Rosto e Metadados Identificados:",
            font=ctk.CTkFont(size=13, weight="bold"),
            text_color=("#1E293B", "#F8FAFC")
        )
        self.lbl_meta_heading.pack(anchor="w", padx=20, pady=(12, 10))

        # Grelha de Cartões / Informações
        self.meta_grid = ctk.CTkFrame(self.meta_card, fg_color="transparent")
        self.meta_grid.pack(fill="x", padx=16, pady=(0, 12))
        self.meta_grid.grid_columnconfigure(1, weight=1)

        self.meta_widgets = {}
        fields_config = [
            ("📘 Título:", "titulo", ("#EEF2FF", "#312E81"), ("#4338CA", "#A5B4FC")),
            ("👥 Autor(es):", "autores", ("#F5F3FF", "#4C1D95"), ("#6D28D9", "#C4B5FD")),
            ("🏛️ Instituição:", "instituicao", ("#ECFDF5", "#064E3B"), ("#047857", "#6EE7B7")),
            ("🎓 Curso:", "curso", ("#F0FDFA", "#134E4A"), ("#0F766E", "#5EEAD4")),
            ("📚 Disciplina:", "disciplina", ("#FEF3C7", "#78350F"), ("#B45309", "#FCD34D")),
            ("👨‍🏫 Orientador:", "docente", ("#F3E8FF", "#581C87"), ("#7E22CE", "#D8B4FE")),
            ("📅 Datação:", "data", ("#FFF1F2", "#881337"), ("#E11D48", "#FDA4AF"))
        ]

        for r, (label_txt, key, bg_badge, fg_badge) in enumerate(fields_config):
            # Badge de categoria
            badge = ctk.CTkLabel(
                self.meta_grid,
                text=label_txt,
                font=ctk.CTkFont(size=11, weight="bold"),
                fg_color=bg_badge,
                text_color=fg_badge,
                corner_radius=6,
                padx=8,
                pady=2,
                width=140,
                anchor="w"
            )
            badge.grid(row=r, column=0, sticky="w", pady=3, padx=(0, 10))

            # Texto do valor
            lbl_val = ctk.CTkLabel(
                self.meta_grid,
                text="—",
                font=ctk.CTkFont(size=11),
                text_color=("#475569", "#94A3B8"),
                anchor="w",
                justify="left"
            )
            lbl_val.grid(row=r, column=1, sticky="w", pady=3)
            self.meta_widgets[key] = lbl_val

        # 4. Consola de Auditoria (Design Estilo Terminal Moderno)
        self.console_card = ctk.CTkFrame(
            self.scroll_container,
            corner_radius=16,
            fg_color=("#0F172A", "#0B0F17"),
            border_width=1,
            border_color=("#334155", "#1E293B")
        )
        self.console_card.pack(fill="x", pady=8)

        self.lbl_console = ctk.CTkLabel(
            self.console_card,
            text="Trilha de Auditoria em Tempo Real:",
            font=ctk.CTkFont(size=11, weight="bold"),
            text_color="#94A3B8"
        )
        self.lbl_console.pack(anchor="w", padx=16, pady=(10, 4))

        self.txt_console = ctk.CTkTextbox(
            self.console_card,
            font=ctk.CTkFont(family="Consolas", size=10),
            fg_color=("#1E293B", "#0F172A"),
            text_color="#38BDF8",
            height=110,
            corner_radius=8
        )
        self.txt_console.pack(fill="x", padx=14, pady=(0, 12))
        self.txt_console.insert("end", "[Pronto] Carregue um documento Word para auditar e compilar.\n")

        # 5. Barra de Progresso e Ações Finais
        self.progress_bar = ctk.CTkProgressBar(
            self.scroll_container,
            corner_radius=8,
            progress_color=("#10B981", "#34D399"),
            fg_color=("#E2E8F0", "#334155")
        )
        self.progress_bar.set(0)
        self.progress_bar.pack(fill="x", pady=(8, 12))

        self.action_box = ctk.CTkFrame(self.scroll_container, fg_color="transparent")
        self.action_box.pack(fill="x", pady=(0, 15))

        self.btn_run = ctk.CTkButton(
            self.action_box,
            text="⚡ Normalizar e Compilar PDF Final",
            height=44,
            font=ctk.CTkFont(size=13, weight="bold"),
            fg_color=("#10B981", "#059669"),
            hover_color=("#059669", "#047857"),
            corner_radius=12,
            command=self._start_processing
        )
        self.btn_run.pack(side="left", fill="x", expand=True, padx=(0, 10))

        self.btn_open_pdf = ctk.CTkButton(
            self.action_box,
            text="Abrir PDF Final",
            height=44,
            fg_color=("#3B8ED0", "#2563EB"),
            hover_color=("#1D4ED8", "#1E40AF"),
            font=ctk.CTkFont(size=13, weight="bold"),
            corner_radius=12,
            state="disabled",
            command=self._open_pdf
        )
        self.btn_open_pdf.pack(side="right", padx=(10, 0))

    def _log(self, message):
        self.txt_console.insert("end", f"[*] {message}\n")
        self.txt_console.see("end")

    def _load_document(self, file_path):
        self.selected_file_path = os.path.abspath(file_path)
        file_name = os.path.basename(self.selected_file_path)
        self.lbl_file_status.configure(
            text=f"Ficheiro carregado: {file_name}",
            text_color=("#4F46E5", "#818CF8")
        )
        self._log(f"Ficheiro selecionado: {file_name}")

        meta = extract_academic_metadata(self.selected_file_path)
        if meta:
            for k, val in meta.items():
                if k in self.meta_widgets:
                    display_text = (val[:80] + "...") if len(val) > 80 else val
                    self.meta_widgets[k].configure(
                        text=display_text,
                        text_color=("#0F172A", "#F8FAFC")
                    )
            self._log("Metadados da folha de rosto organizados e validados.")

    def _select_file_dialog(self):
        path = filedialog.askopenfilename(
            title="Escolha o ficheiro Word",
            filetypes=[("Documentos Word", "*.docx")]
        )
        if path:
            self._load_document(path)

    def _on_file_drop(self, event):
        raw_data = event.data
        if raw_data:
            files = self.tk.splitlist(raw_data)
            if files and files[0].lower().endswith(".docx"):
                self._load_document(files[0])
            else:
                messagebox.showwarning("Formato Inválido", "Por favor, arraste apenas ficheiros .docx.")

    def _start_processing(self):
        if not self.selected_file_path or not os.path.exists(self.selected_file_path):
            messagebox.showwarning("Aviso", "Selecione primeiro um documento Word (.docx) válido.")
            return

        self.btn_run.configure(state="disabled")
        self.btn_open_pdf.configure(state="disabled")
        self.progress_bar.set(0)

        worker = threading.Thread(target=self._process_pipeline, daemon=True)
        worker.start()

    def _process_pipeline(self):
        import pythoncom
        pythoncom.CoInitialize()

        try:
            base_dir = os.path.dirname(self.selected_file_path)
            normalized_docx = os.path.join(base_dir, "Qualidade_Normalizado.docx")
            self.output_pdf_path = os.path.join(base_dir, "Qualidade_de_Sistemas_e_Software_Final.pdf")

            self._log("Iniciando auditoria determinística na árvore OpenXML...")
            self.progress_bar.set(0.2)
            time.sleep(0.3)

            # Executa o motor configurado
            try:
                import apa_engine_v2
                config_file = os.path.join(os.path.dirname(__file__), "apa_config.json")
                if not os.path.exists(config_file):
                    config_file = "apa_config.json"
                apa_engine_v2.process_document(self.selected_file_path, normalized_docx, config_file)
            except ImportError:
                import apa_docx_engine
                apa_docx_engine.process_document(self.selected_file_path, normalized_docx)

            self._log("Hierarquia de Sumário, listas e referências normalizadas.")
            self.progress_bar.set(0.6)
            time.sleep(0.3)

            self._log("Iniciando automação do Microsoft Word via Win32 COM...")
            self.progress_bar.set(0.8)

            try:
                import apa_compiler_v2
                apa_compiler_v2.compile_pdf(normalized_docx, self.output_pdf_path)
            except ImportError:
                import docx_pdf_compiler
                docx_pdf_compiler.compile_pdf(normalized_docx, self.output_pdf_path)

            self._log(f"PDF vetorial compilado: {os.path.basename(self.output_pdf_path)}")
            self.progress_bar.set(1.0)
            self.after(0, self._on_success)

        except Exception as err:
            self._log(f"FALHA: {err}")
            self.after(0, lambda: messagebox.showerror("Erro de Processamento", str(err)))
            self.after(0, lambda: self.btn_run.configure(state="normal"))
        finally:
            pythoncom.CoUninitialize()

    def _on_success(self):
        self.btn_run.configure(state="normal")
        self.btn_open_pdf.configure(state="normal")
        messagebox.showinfo("Sucesso", "Documento normalizado e PDF gerado com sucesso!")

    def _open_pdf(self):
        if self.output_pdf_path and os.path.exists(self.output_pdf_path):
            os.startfile(self.output_pdf_path)
        else:
            messagebox.showwarning("Aviso", "O ficheiro PDF ainda não se encontra disponível.")

if __name__ == "__main__":
    app = ApaDesktopApp()
    app.mainloop()