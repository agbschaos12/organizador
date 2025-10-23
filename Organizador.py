# Organizador.py
# Dependências opcionais:
#  - customtkinter (recomendado para aparência)
#  - tkcalendar (opcional para calendário visual)
# Se preferir não instalar tkcalendar, os campos de data continuam funcionando por entrada manual.

import os
import sys
import json
import shutil
import logging
import threading
import queue
from datetime import datetime
from pathlib import Path
from typing import List, Dict, Optional, Callable

import customtkinter
import tkinter
from tkinter import filedialog, messagebox

# Tentativa de usar tkcalendar; se não houver, o app continua com entradas manuais e mostra instrução
try:
    from tkcalendar import Calendar
    HAS_TKCALENDAR = True
except Exception:
    HAS_TKCALENDAR = False

# ------------------- file ops -------------------
LOGFILE = "organizer.log"
UNDO_FILE = "undo_record.json"

# Logger setup
logger = logging.getLogger("organizer")
if not logger.handlers:
    fh = logging.FileHandler(LOGFILE, encoding="utf-8")
    fmt = logging.Formatter("%(asctime)s - %(levelname)s - %(message)s")
    fh.setFormatter(fmt)
    logger.addHandler(fh)
logger.setLevel(logging.INFO)

def _parse_date(date_str: Optional[str]) -> Optional[datetime]:
    if not date_str:
        return None
    return datetime.strptime(date_str, "%Y-%m-%d")

def get_matched_files(org: Dict) -> List[Path]:
    """
    Retorna lista de Path de arquivos que batem com as keywords e filtro de data.
    Espera org com chaves:
      - source (str)
      - keywords (List[str])
      - date_filter_enabled (bool)
      - start_date (str or None)
      - end_date (str or None)
    """
    source = Path(org["source"])
    keywords = [k.lower() for k in org.get("keywords", [])]
    if not source.is_dir():
        logger.warning("Source não é diretório: %s", source)
        return []

    start = _parse_date(org.get("start_date")) if org.get("date_filter_enabled") else None
    end = _parse_date(org.get("end_date")) if org.get("date_filter_enabled") else None
    if end:
        end = end.replace(hour=23, minute=59, second=59)

    matched: List[Path] = []
    for p in source.iterdir():
        if p.is_dir():
            continue
        try:
            mtime = datetime.fromtimestamp(p.stat().st_mtime)
        except Exception:
            logger.exception("Erro lendo metadados de %s", p)
            continue
        if start and mtime < start:
            continue
        if end and mtime > end:
            continue
        name = p.name.lower()
        for kw in keywords:
            if kw and kw in name:
                matched.append(p)
                break

    logger.info("Encontrados %d arquivo(s) em %s", len(matched), source)
    return matched

def _unique_dest(dest_dir: Path, filename: str) -> Path:
    dest_path = dest_dir / filename
    if not dest_path.exists():
        return dest_path
    base = dest_path.stem
    ext = dest_path.suffix
    counter = 1
    while True:
        candidate = dest_dir / f"{base} ({counter}){ext}"
        if not candidate.exists():
            return candidate
        counter += 1

def move_files(files: List[Path],
               destination: str,
               dry_run: bool = False,
               progress_callback: Optional[Callable[[int, int], None]] = None
               ) -> List[Dict]:
    """
    Move os arquivos para destination. Se dry_run=True não move, apenas simula.
    progress_callback(completed, total) é chamado após cada tentativa de arquivo.
    Retorna lista de dicts: {"source": str, "dest": str, "action": "moved"|"would_move"|"error"}
    """
    dest_dir = Path(destination)
    try:
        dest_dir.mkdir(parents=True, exist_ok=True)
    except Exception as e:
        logger.exception("Não foi possível criar/verificar destino %s: %s", dest_dir, e)
        raise

    results: List[Dict] = []
    total = len(files)
    completed = 0

    for src in files:
        try:
            dest_path = _unique_dest(dest_dir, src.name)
            if dry_run:
                action = "would_move"
                logger.info("[DRY] %s -> %s", src, dest_path)
            else:
                shutil.move(str(src), str(dest_path))
                action = "moved"
                logger.info("Moved %s -> %s", src, dest_path)
            results.append({"source": str(src), "dest": str(dest_path), "action": action})
        except Exception as e:
            logger.exception("Erro movendo %s para %s: %s", src, destination, e)
            results.append({"source": str(src), "dest": str(destination), "action": f"error: {e}"})
        completed += 1
        if progress_callback:
            try:
                progress_callback(completed, total)
            except Exception:
                logger.exception("Erro no progress_callback")
    return results

def save_undo_record(record: Dict):
    try:
        with open(UNDO_FILE, "w", encoding="utf-8") as f:
            json.dump(record, f, ensure_ascii=False, indent=2)
        logger.info("Undo record salvo (%d operações)", len(record.get("operations", [])))
    except Exception:
        logger.exception("Erro salvando undo record")

def load_undo_record() -> Optional[Dict]:
    try:
        p = Path(UNDO_FILE)
        if not p.exists():
            return None
        with p.open("r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        logger.exception("Erro carregando undo record")
        return None

def undo_last(progress_callback: Optional[Callable[[int, int], None]] = None) -> Dict:
    """
    Tenta desfazer a última operação gravada em UNDO_FILE.
    Retorna um dict com resultado e detalhes.
    """
    record = load_undo_record()
    if not record:
        return {"ok": False, "error": "Nenhum registro de undo encontrado."}

    ops = record.get("operations", [])
    total = len(ops)
    completed = 0
    results = []
    for op in reversed(ops):
        src = Path(op["dest"])
        dest = Path(op["source"])
        try:
            if not src.exists():
                results.append({"source": str(src), "dest": str(dest), "action": "skipped_not_found"})
            else:
                final_dest = _unique_dest(dest.parent, dest.name)
                shutil.move(str(src), str(final_dest))
                results.append({"source": str(src), "dest": str(final_dest), "action": "restored"})
        except Exception as e:
            logger.exception("Erro desfazendo %s -> %s: %s", src, dest, e)
            results.append({"source": str(src), "dest": str(dest), "action": f"error: {e}"})
        completed += 1
        if progress_callback:
            try:
                progress_callback(completed, total)
            except Exception:
                logger.exception("Erro no progress_callback durante undo")

    try:
        Path(UNDO_FILE).unlink(missing_ok=True)
    except Exception:
        logger.exception("Erro removendo UNDO_FILE")

    return {"ok": True, "results": results}

# Expor logger compatível com o código GUI (apenas para nomes anteriores)
file_logger = logger

# ------------------- Fim file_ops -------------------

# Aparência e tema
customtkinter.set_appearance_mode("Dark")
customtkinter.set_default_color_theme("dark-blue")

CONFIG_FILE = "config.json"

class App(customtkinter.CTk):
    def __init__(self):
        super().__init__()
        self.title("Organizador de Arquivos")
        self.geometry("820x640")
        self.minsize(760, 520)

        # estado
        self.organizations = self.load_organizations()
        self.editing_org_original_name = None
        self.last_execution_record = None  # dict retornada após execução (apenas em memória)
        self._thread_queue = queue.Queue()

        # Layout principal
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.main_frame = customtkinter.CTkFrame(self, corner_radius=8)
        self.main_frame.grid(row=0, column=0, sticky="nsew", padx=16, pady=16)
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(6, weight=1)

        # Header
        header = customtkinter.CTkLabel(self.main_frame, text="Organizador de Arquivos", font=customtkinter.CTkFont(size=22, weight="bold"))
        header.grid(row=0, column=0, padx=12, pady=(12,6), sticky="w")
        sub = customtkinter.CTkLabel(self.main_frame, text="Selecione, visualize, edite, duplique, exporte, importe ou execute organizações.", font=customtkinter.CTkFont(size=12))
        sub.grid(row=1, column=0, padx=12, pady=(0,12), sticky="w")

        # selection row
        sel_frame = customtkinter.CTkFrame(self.main_frame, fg_color="transparent")
        sel_frame.grid(row=2, column=0, sticky="ew", padx=12, pady=(0,12))
        sel_frame.grid_columnconfigure(1, weight=1)

        org_names = list(self.organizations.keys()) or ["Nenhuma organização criada"]
        self.org_combobox = customtkinter.CTkComboBox(sel_frame, values=org_names, width=420)
        self.org_combobox.grid(row=0, column=0, padx=(0,8), pady=4, sticky="w")
        self.org_combobox.set(org_names[0])

        btns = customtkinter.CTkFrame(sel_frame, fg_color="transparent")
        btns.grid(row=0, column=1, sticky="e")
        self.run_button = customtkinter.CTkButton(btns, text="Executar", width=100, command=self.preview_or_run)
        self.run_button.grid(row=0, column=0, padx=(0,6))
        self.preview_button = customtkinter.CTkButton(btns, text="Pré-visualizar", width=110, command=self.show_preview)
        self.preview_button.grid(row=0, column=1, padx=(0,6))
        self.details_button = customtkinter.CTkButton(btns, text="Ver Detalhes", width=110, command=self.show_details)
        self.details_button.grid(row=0, column=2, padx=(0,6))
        self.edit_button = customtkinter.CTkButton(btns, text="Editar", width=90, command=self.edit_organization)
        self.edit_button.grid(row=0, column=3, padx=(0,6))
        self.duplicate_button = customtkinter.CTkButton(btns, text="Duplicar", width=90, command=self.duplicate_organization)
        self.duplicate_button.grid(row=0, column=4, padx=(0,6))
        self.delete_button = customtkinter.CTkButton(btns, text="Excluir", width=90, fg_color="#ff5555", hover_color="#ff7777", command=self.delete_organization)
        self.delete_button.grid(row=0, column=5, padx=(0,0))

        # separator
        sep = customtkinter.CTkFrame(self.main_frame, height=2, fg_color="#3a3a3a")
        sep.grid(row=3, column=0, sticky="ew", padx=12, pady=(12,12))

        # details area
        details_container = customtkinter.CTkFrame(self.main_frame, corner_radius=6)
        details_container.grid(row=4, column=0, sticky="nsew", padx=12, pady=(0,12))
        details_container.grid_columnconfigure(0, weight=1)
        details_container.grid_rowconfigure(0, weight=1)
        self.details_label = customtkinter.CTkLabel(details_container, text="Detalhes aparecerão aqui.", anchor="w", justify="left", wraplength=760)
        self.details_label.grid(row=0, column=0, sticky="nwes", padx=12, pady=12)

        # footer actions
        footer = customtkinter.CTkFrame(self.main_frame, fg_color="transparent")
        footer.grid(row=5, column=0, sticky="ew", padx=12, pady=(0,6))
        footer.grid_columnconfigure(0, weight=1)
        self.create_button = customtkinter.CTkButton(footer, text="Criar Nova Organização", command=self.show_creation_frame)
        self.create_button.grid(row=0, column=0, sticky="w")
        self.import_button = customtkinter.CTkButton(footer, text="Importar JSON", command=self.import_organization)
        self.import_button.grid(row=0, column=1, sticky="e")
        self.export_button = customtkinter.CTkButton(footer, text="Exportar", command=self.export_organization)
        self.export_button.grid(row=0, column=2, sticky="e", padx=(6,0))
        self.undo_button = customtkinter.CTkButton(footer, text="Desfazer Última", command=self.undo_last_execution)
        self.undo_button.grid(row=0, column=3, sticky="e", padx=(6,0))
        self.open_logs_button = customtkinter.CTkButton(footer, text="Abrir Logs", command=self.open_logs)
        self.open_logs_button.grid(row=0, column=4, sticky="e", padx=(6,0))

        self.org_count_label = customtkinter.CTkLabel(self.main_frame, text=f"{len(self.organizations)} organização(ões) salvas")
        self.org_count_label.grid(row=6, column=0, sticky="w", padx=12, pady=(0,12))

        # Creation frame (hidden initially)
        self.creation_frame = customtkinter.CTkFrame(self, corner_radius=8)
        self.creation_frame.grid_columnconfigure(1, weight=1)

        # Creation widgets
        self.creation_label = customtkinter.CTkLabel(self.creation_frame, text="Criar / Editar Organização", font=customtkinter.CTkFont(size=20, weight="bold"))
        self.creation_label.grid(row=0, column=0, columnspan=3, padx=20, pady=(20,10))

        self.name_label = customtkinter.CTkLabel(self.creation_frame, text="Nome:")
        self.name_label.grid(row=1, column=0, padx=(20,5), pady=5, sticky="w")
        self.name_entry = customtkinter.CTkEntry(self.creation_frame, placeholder_text="Ex: Relatórios")
        self.name_entry.grid(row=1, column=1, columnspan=2, padx=(0,20), pady=5, sticky="ew")

        self.source_label = customtkinter.CTkLabel(self.creation_frame, text="Pasta de Origem:")
        self.source_label.grid(row=2, column=0, padx=(20,5), pady=5, sticky="w")
        self.source_entry = customtkinter.CTkEntry(self.creation_frame, placeholder_text="Selecione a pasta onde os arquivos estão")
        self.source_entry.grid(row=2, column=1, padx=0, pady=5, sticky="ew")
        self.source_btn = customtkinter.CTkButton(self.creation_frame, text="Procurar...", width=80, command=lambda: self.select_directory(self.source_entry))
        self.source_btn.grid(row=2, column=2, padx=(5,20), pady=5)

        self.keywords_label = customtkinter.CTkLabel(self.creation_frame, text="Palavras-chave:")
        self.keywords_label.grid(row=3, column=0, padx=(20,5), pady=5, sticky="w")
        self.keywords_entry = customtkinter.CTkEntry(self.creation_frame, placeholder_text="Use vírgulas para separar (fatura, nota)")
        self.keywords_entry.grid(row=3, column=1, columnspan=2, padx=(0,20), pady=5, sticky="ew")

        self.dest_label = customtkinter.CTkLabel(self.creation_frame, text="Pasta de Destino:")
        self.dest_label.grid(row=4, column=0, padx=(20,5), pady=5, sticky="w")
        self.dest_entry = customtkinter.CTkEntry(self.creation_frame, placeholder_text="Selecione a pasta destino")
        self.dest_entry.grid(row=4, column=1, padx=0, pady=5, sticky="ew")
        self.dest_btn = customtkinter.CTkButton(self.creation_frame, text="Procurar...", width=80, command=lambda: self.select_directory(self.dest_entry))
        self.dest_btn.grid(row=4, column=2, padx=(5,20), pady=5)

        # Date filter
        self.date_filter_checkbox = customtkinter.CTkCheckBox(self.creation_frame, text="Habilitar filtro por data", command=self.toggle_date_fields)
        self.date_filter_checkbox.grid(row=5, column=0, columnspan=3, padx=20, pady=(20,5), sticky="w")

        self.date_frame = customtkinter.CTkFrame(self.creation_frame, fg_color="transparent")
        self.date_frame.grid(row=6, column=0, columnspan=3, padx=20, pady=0, sticky="ew")
        self.date_frame.grid_columnconfigure(1, weight=1)

        # Start date entry + calendar button
        self.start_label = customtkinter.CTkLabel(self.date_frame, text="Data de Início:")
        self.start_label.grid(row=0, column=0, padx=(0,5), pady=5, sticky="w")
        self.start_entry = customtkinter.CTkEntry(self.date_frame, placeholder_text="AAAA-MM-DD")
        self.start_entry.grid(row=0, column=1, padx=(0,5), pady=5, sticky="ew")
        if HAS_TKCALENDAR:
            self.start_cal_btn = customtkinter.CTkButton(self.date_frame, text="Selecionar...", width=100, command=lambda: self.open_calendar(self.start_entry))
            self.start_cal_btn.grid(row=0, column=2, padx=(5,20), pady=5)
        else:
            self.start_cal_btn = customtkinter.CTkButton(self.date_frame, text="Selecionar...", width=100, command=self._no_calendar_installed)
            self.start_cal_btn.grid(row=0, column=2, padx=(5,20), pady=5)

        # End date entry + calendar button
        self.end_label = customtkinter.CTkLabel(self.date_frame, text="Data de Fim:")
        self.end_label.grid(row=1, column=0, padx=(0,5), pady=5, sticky="w")
        self.end_entry = customtkinter.CTkEntry(self.date_frame, placeholder_text="AAAA-MM-DD")
        self.end_entry.grid(row=1, column=1, padx=(0,5), pady=5, sticky="ew")
        if HAS_TKCALENDAR:
            self.end_cal_btn = customtkinter.CTkButton(self.date_frame, text="Selecionar...", width=100, command=lambda: self.open_calendar(self.end_entry))
            self.end_cal_btn.grid(row=1, column=2, padx=(5,20), pady=5)
        else:
            self.end_cal_btn = customtkinter.CTkButton(self.date_frame, text="Selecionar...", width=100, command=self._no_calendar_installed)
            self.end_cal_btn.grid(row=1, column=2, padx=(5,20), pady=5)

        # Save / cancel
        self.save_btn = customtkinter.CTkButton(self.creation_frame, text="Salvar Organização", command=self.save_organization)
        self.save_btn.grid(row=7, column=1, padx=5, pady=20, sticky="ew")
        self.cancel_btn = customtkinter.CTkButton(self.creation_frame, text="Voltar", fg_color="transparent", border_width=2, command=self.show_main_frame)
        self.cancel_btn.grid(row=7, column=2, padx=5, pady=20, sticky="ew")

        self.toggle_date_fields()

        # menu and shortcuts
        self.create_menu()

        # checar fila de threads
        self.after(200, self._process_thread_queue)

    # ----------------- Calendar popup -----------------
    def _no_calendar_installed(self):
        messagebox.showinfo(
            "Calendário opcional",
            "Para usar o calendário visual, instale a biblioteca 'tkcalendar'.\n\n"
            "Comando:\n    pip install tkcalendar\n\n"
            "Enquanto isso, você pode digitar a data no formato AAAA-MM-DD."
        )

    def open_calendar(self, entry_widget):
        """
        Abre um popup com um calendário (tkcalendar.Calendar). Ao selecionar uma data,
        ela é inserida no entry_widget no formato YYYY-MM-DD.
        """
        if not HAS_TKCALENDAR:
            self._no_calendar_installed()
            return

        # Try to parse existing value to set initial date on calendar
        init_date = None
        try:
            val = entry_widget.get().strip()
            if val:
                init_date = datetime.strptime(val, "%Y-%m-%d").date()
        except Exception:
            init_date = None

        popup = customtkinter.CTkToplevel(self)
        popup.title("Selecione a data")
        try:
            popup.transient(self)
        except Exception:
            pass

        popup.geometry("+%d+%d" % (self.winfo_rootx() + 100, self.winfo_rooty() + 100))

        cal = Calendar(popup, selectmode="day", date_pattern="yyyy-mm-dd")
        if init_date:
            try:
                cal.selection_set(init_date)
            except Exception:
                pass
        cal.pack(padx=8, pady=8, expand=True, fill="both")

        btn_frame = customtkinter.CTkFrame(popup, fg_color="transparent")
        btn_frame.pack(fill="x", padx=8, pady=(0,8))

        def on_ok():
            try:
                sel = cal.selection_get()  # returns datetime.date
                entry_widget.delete(0, "end")
                entry_widget.insert(0, sel.strftime("%Y-%m-%d"))
            except Exception:
                try:
                    s = cal.get_date()
                    entry_widget.delete(0, "end")
                    entry_widget.insert(0, s)
                except Exception:
                    pass
            popup.destroy()

        def on_cancel():
            popup.destroy()

        ok_btn = customtkinter.CTkButton(btn_frame, text="OK", command=on_ok)
        ok_btn.pack(side="right", padx=(6,0))
        cancel_btn = customtkinter.CTkButton(btn_frame, text="Cancelar", command=on_cancel)
        cancel_btn.pack(side="right", padx=(0,6))

        try:
            popup.grab_set()
        except Exception:
            pass

    # ----------------- Config / IO -----------------
    def load_organizations(self):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}

    def save_organizations_to_file(self):
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(self.organizations, f, indent=4, ensure_ascii=False)
        except Exception as e:
            messagebox.showerror("Erro ao salvar", f"Não foi possível salvar config:\n{e}")

    def update_combobox(self):
        names = list(self.organizations.keys())
        if not names:
            names = ["Nenhuma organização criada"]
        self.org_combobox.configure(values=names)
        self.org_combobox.set(names[0])
        self.org_count_label.configure(text=f"{len(self.organizations)} organização(ões) salvas")
        self.details_label.configure(text="Detalhes aparecerão aqui.")

    # ----------------- UI helpers -----------------
    def show_creation_frame(self):
        self.editing_org_original_name = None
        self.name_entry.delete(0, "end")
        self.source_entry.delete(0, "end")
        self.keywords_entry.delete(0, "end")
        self.dest_entry.delete(0, "end")
        self.start_entry.delete(0, "end")
        self.end_entry.delete(0, "end")
        self.date_filter_checkbox.deselect()
        self.toggle_date_fields()
        self.main_frame.grid_forget()
        self.creation_frame.grid(row=0, column=0, sticky="nsew", padx=16, pady=16)

    def show_main_frame(self):
        self.editing_org_original_name = None
        self.creation_frame.grid_forget()
        self.main_frame.grid(row=0, column=0, sticky="nsew", padx=16, pady=16)
        self.update_combobox()

    def toggle_date_fields(self):
        if self.date_filter_checkbox.get() == 1:
            self.date_frame.grid(row=6, column=0, columnspan=3, padx=20, pady=0, sticky="ew")
        else:
            self.date_frame.grid_forget()

    def select_directory(self, entry_widget):
        path = filedialog.askdirectory()
        if path:
            entry_widget.delete(0, "end")
            entry_widget.insert(0, path)

    # ----------------- Create / Edit / Delete / Duplicate / Import / Export -----------------
    def save_organization(self):
        name = self.name_entry.get().strip()
        source = self.source_entry.get().strip()
        keywords = self.keywords_entry.get().strip()
        dest = self.dest_entry.get().strip()
        date_enabled = self.date_filter_checkbox.get() == 1

        start = self.start_entry.get().strip()
        end = self.end_entry.get().strip()

        if not all([name, source, keywords, dest]):
            messagebox.showerror("Erro", "Nome, origem, palavras-chave e destino são obrigatórios.")
            return

        if date_enabled:
            for d in [start, end]:
                if d:
                    try:
                        datetime.strptime(d, "%Y-%m-%d")
                    except ValueError:
                        messagebox.showerror("Erro de Formato", f"A data '{d}' não está no formato AAAA-MM-DD.")
                        return

        # salvar/atualizar (tratamento de renomeação)
        if self.editing_org_original_name:
            original = self.editing_org_original_name
            if name != original and name in self.organizations:
                if not messagebox.askyesno("Confirmar", f"Já existe organização '{name}'. Sobrescrever?"):
                    return
            # remover original se renomeado
            if name != original:
                self.organizations.pop(original, None)
        else:
            if name in self.organizations:
                if not messagebox.askyesno("Confirmar", f"Já existe organização '{name}'. Sobrescrever?"):
                    return

        self.organizations[name] = {
            "source": source,
            "keywords": [k.strip().lower() for k in keywords.split(",") if k.strip()],
            "destination": dest,
            "date_filter_enabled": date_enabled,
            "start_date": start or "",
            "end_date": end or ""
        }
        self.save_organizations_to_file()
        messagebox.showinfo("Sucesso", f"Organização '{name}' salva.")
        self.show_main_frame()

    def delete_organization(self):
        org_name = self.org_combobox.get()
        if not org_name or org_name == "Nenhuma organização criada":
            messagebox.showwarning("Aviso", "Selecione uma organização válida.")
            return
        if org_name not in self.organizations:
            messagebox.showerror("Erro", "Organização não encontrada.")
            return
        if not messagebox.askyesno("Confirmar Exclusão", f"Excluir '{org_name}'?"):
            return
        self.organizations.pop(org_name, None)
        self.save_organizations_to_file()
        messagebox.showinfo("Removido", f"Organização '{org_name}' excluída.")
        self.update_combobox()

    def edit_organization(self):
        org_name = self.org_combobox.get()
        if not org_name or org_name == "Nenhuma organização criada":
            messagebox.showwarning("Aviso", "Selecione uma organização válida.")
            return
        org = self.organizations.get(org_name)
        if not org:
            messagebox.showerror("Erro", "Organização não encontrada.")
            return
        self.editing_org_original_name = org_name
        self.name_entry.delete(0, "end")
        self.name_entry.insert(0, org_name)
        self.source_entry.delete(0, "end")
        self.source_entry.insert(0, org.get("source", ""))
        self.keywords_entry.delete(0, "end")
        self.keywords_entry.insert(0, ", ".join(org.get("keywords", [])))
        self.dest_entry.delete(0, "end")
        self.dest_entry.insert(0, org.get("destination", ""))
        if org.get("date_filter_enabled"):
            self.date_filter_checkbox.select()
            # fill entries as strings (YYYY-MM-DD)
            self.start_entry.delete(0, "end")
            if org.get("start_date"):
                self.start_entry.insert(0, org.get("start_date"))
            self.end_entry.delete(0, "end")
            if org.get("end_date"):
                self.end_entry.insert(0, org.get("end_date"))
        else:
            self.date_filter_checkbox.deselect()
            self.start_entry.delete(0, "end")
            self.end_entry.delete(0, "end")
        self.toggle_date_fields()
        self.main_frame.grid_forget()
        self.creation_frame.grid(row=0, column=0, sticky="nsew", padx=16, pady=16)

    def duplicate_organization(self):
        org_name = self.org_combobox.get()
        if org_name not in self.organizations or org_name == "Nenhuma organização criada":
            messagebox.showwarning("Aviso", "Selecione uma organização válida.")
            return
        base_name = f"{org_name} (cópia)"
        i = 1
        new_name = base_name
        while new_name in self.organizations:
            i += 1
            new_name = f"{base_name} {i}"
        self.organizations[new_name] = dict(self.organizations[org_name])
        self.save_organizations_to_file()
        messagebox.showinfo("Duplicado", f"Organização duplicada como '{new_name}'.")
        self.update_combobox()

    def export_organization(self):
        org_name = self.org_combobox.get()
        if org_name not in self.organizations or org_name == "Nenhuma organização criada":
            messagebox.showwarning("Aviso", "Selecione uma organização válida.")
            return
        org = self.organizations[org_name]
        path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON", "*.json")], title="Exportar organização como...")
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8") as f:
                json.dump({org_name: org}, f, indent=4, ensure_ascii=False)
            messagebox.showinfo("Exportado", f"Organização exportada para {path}")
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível exportar:\n{e}")

    def import_organization(self):
        path = filedialog.askopenfilename(title="Importar organização (JSON)", filetypes=[("JSON", "*.json")])
        if not path:
            return
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            # esperar estrutura { "Nome": { ... } } ou múltiplas entradas
            added = 0
            for k, v in data.items():
                if k in self.organizations:
                    if not messagebox.askyesno("Confirmar", f"Organização '{k}' existe. Sobrescrever?"):
                        continue
                self.organizations[k] = v
                added += 1
            self.save_organizations_to_file()
            messagebox.showinfo("Importado", f"{added} organização(ões) importada(s).")
            self.update_combobox()
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível importar:\n{e}")

    # ----------------- Preview and Execution (threads + progress) -----------------
    def preview_or_run(self):
        # shortcut: open preview first
        self.show_preview()

    def show_preview(self):
        org_name = self.org_combobox.get()
        if not org_name or org_name == "Nenhuma organização criada":
            messagebox.showwarning("Aviso", "Selecione uma organização válida.")
            return
        org = self.organizations.get(org_name)
        if not org:
            messagebox.showerror("Erro", "Organização não encontrada.")
            return

        matched = get_matched_files(org)
        preview_win = customtkinter.CTkToplevel(self)
        preview_win.title(f"Pré-visualizar: {org_name}")
        preview_win.geometry("760x480")

        info = customtkinter.CTkLabel(preview_win, text=f"{len(matched)} arquivo(s) encontrados que correspondem aos filtros.")
        info.pack(padx=12, pady=(12,6), anchor="w")

        frame = customtkinter.CTkFrame(preview_win)
        frame.pack(fill="both", expand=True, padx=12, pady=(0,12))
        scrollbar = customtkinter.CTkScrollbar(frame, orientation="vertical")
        scrollbar.pack(side="right", fill="y")
        lb = tkinter.Listbox(frame, yscrollcommand=scrollbar.set)
        for p in matched:
            lb.insert("end", str(p))
        lb.pack(side="left", fill="both", expand=True)
        scrollbar.configure(command=lb.yview)

        btn_frame = customtkinter.CTkFrame(preview_win, fg_color="transparent")
        btn_frame.pack(fill="x", padx=12, pady=(0,12))
        exec_btn = customtkinter.CTkButton(btn_frame, text="Executar Movimento", fg_color="#ff7744",
                                          command=lambda: self._confirm_and_execute_preview(preview_win, matched, org))
        exec_btn.pack(side="right", padx=(6,0))
        close_btn = customtkinter.CTkButton(btn_frame, text="Fechar", command=preview_win.destroy)
        close_btn.pack(side="right", padx=(0,6))

    def _confirm_and_execute_preview(self, preview_win, matched_files, org):
        if not matched_files:
            messagebox.showinfo("Nada a Fazer", "Não há arquivos correspondentes.")
            return
        if not messagebox.askyesno("Confirmar Execução", f"Executar e mover {len(matched_files)} arquivo(s)?"):
            return
        preview_win.destroy()
        # cria janela de progresso e executa em thread
        self._run_move_in_thread(matched_files, org, dry_run=False)

    def _run_move_in_thread(self, matched_files, org, dry_run=False):
        # progress window
        pw = customtkinter.CTkToplevel(self)
        pw.title("Executando organização")
        pw.geometry("520x160")
        label = customtkinter.CTkLabel(pw, text="Movendo arquivos...")
        label.pack(padx=12, pady=(12,6))
        progress = customtkinter.CTkProgressBar(pw, orientation="horizontal")
        progress.set(0.0)
        progress.pack(fill="x", padx=12, pady=(6,12))
        status_lbl = customtkinter.CTkLabel(pw, text="0 / 0")
        status_lbl.pack(padx=12, pady=(0,12))

        total = len(matched_files)
        if total == 0:
            messagebox.showinfo("Nada a mover", "Não há arquivos correspondentes.")
            pw.destroy()
            return

        # progress callback enfileira atualizações para o thread principal
        def progress_cb(completed, tot):
            self._thread_queue.put(("progress", completed, tot))

        def worker():
            try:
                results = move_files(matched_files, org["destination"], dry_run=dry_run, progress_callback=progress_cb)
                # após mover, gera undo record (apenas operações efetivas moved)
                record = {
                    "timestamp": datetime.utcnow().isoformat(),
                    "operations": [{"source": r["source"], "dest": r["dest"], "action": r["action"]} for r in results]
                }
                # salvar undo apenas quando não for dry_run and there are moved files
                if not dry_run and any(r.get("action") == "moved" for r in results):
                    save_undo_record(record)
                # store in app state for UI
                self._thread_queue.put(("done", results))
            except Exception as e:
                file_logger.exception("Erro na thread de execução: %s", e)
                self._thread_queue.put(("error", str(e)))

        t = threading.Thread(target=worker, daemon=True)
        t.start()

        # função local para atualizar UI a partir da fila
        def handle_queue():
            try:
                while not self._thread_queue.empty():
                    item = self._thread_queue.get_nowait()
                    if item[0] == "progress":
                        completed, tot = item[1], item[2]
                        progress.set(completed / tot if tot else 0.0)
                        status_lbl.configure(text=f"{completed} / {tot}")
                    elif item[0] == "done":
                        results = item[1]
                        moved = sum(1 for r in results if r.get("action") == "moved")
                        messagebox.showinfo("Concluído", f"Operação finalizada.\n{moved} arquivo(s) movidos.")
                        pw.destroy()
                        self.update_combobox()
                    elif item[0] == "error":
                        messagebox.showerror("Erro na Execução", f"Ocorreu erro: {item[1]}")
                        pw.destroy()
            except Exception:
                file_logger.exception("Erro processando fila de thread")
            finally:
                if pw.winfo_exists():
                    self.after(200, handle_queue)

        # inicia loop de checagem
        self.after(200, handle_queue)

    # ----------------- Undo -----------------
    def undo_last_execution(self):
        # tenta carregar undo record e perguntar confirmação
        rec = load_undo_record()
        if not rec:
            messagebox.showinfo("Nada a Desfazer", "Nenhuma operação anterior encontrada.")
            return
        ops = rec.get("operations", [])
        if not ops:
            messagebox.showinfo("Nada a Desfazer", "Nenhuma operação anterior encontrada.")
            return
        if not messagebox.askyesno("Confirmar Desfazer", f"Desfazer última execução com {len(ops)} operações?"):
            return

        # mostra janela de progresso de undo
        pw = customtkinter.CTkToplevel(self)
        pw.title("Desfazendo...")
        pw.geometry("520x140")
        label = customtkinter.CTkLabel(pw, text="Desfazendo últimas operações...")
        label.pack(padx=12, pady=(12,6))
        progress = customtkinter.CTkProgressBar(pw)
        progress.set(0.0)
        progress.pack(fill="x", padx=12, pady=(6,12))
        status_lbl = customtkinter.CTkLabel(pw, text="0 / 0")
        status_lbl.pack(padx=12, pady=(0,12))

        def progress_cb(c, t):
            self._thread_queue.put(("undo_progress", c, t))

        def worker_undo():
            try:
                result = undo_last(progress_callback=progress_cb)
                self._thread_queue.put(("undo_done", result))
            except Exception as e:
                file_logger.exception("Erro no undo thread: %s", e)
                self._thread_queue.put(("undo_error", str(e)))

        t = threading.Thread(target=worker_undo, daemon=True)
        t.start()

        def handle_q():
            try:
                while not self._thread_queue.empty():
                    item = self._thread_queue.get_nowait()
                    if item[0] == "undo_progress":
                        c, t = item[1], item[2]
                        progress.set(c / t if t else 0.0)
                        status_lbl.configure(text=f"{c} / {t}")
                    elif item[0] == "undo_done":
                        res = item[1]
                        messagebox.showinfo("Desfeito", "Operação de desfazer concluída.")
                        pw.destroy()
                        self.update_combobox()
                    elif item[0] == "undo_error":
                        messagebox.showerror("Erro", f"Erro durante undo: {item[1]}")
                        pw.destroy()
            except Exception:
                file_logger.exception("Erro processando undo fila")
            finally:
                if pw.winfo_exists():
                    self.after(200, handle_q)

        self.after(200, handle_q)

    # ----------------- Details / Logs -----------------
    def show_details(self):
        org_name = self.org_combobox.get()
        if not org_name or org_name == "Nenhuma organização criada":
            messagebox.showwarning("Aviso", "Selecione uma organização válida.")
            return
        org = self.organizations.get(org_name)
        if not org:
            messagebox.showerror("Erro", "Organização não encontrada.")
            return
        keywords = ", ".join(org.get("keywords", []))
        date_filter = "Sim" if org.get("date_filter_enabled") else "Não"
        start = org.get("start_date") or "—"
        end = org.get("end_date") or "—"
        details = (
            f"Nome: {org_name}\n"
            f"Pasta de Origem: {org.get('source')}\n"
            f"Pasta de Destino: {org.get('destination')}\n"
            f"Palavras-chave: {keywords}\n"
            f"Filtro por data: {date_filter}\n"
            f"Data de Início: {start}\n"
            f"Data de Fim: {end}"
        )
        self.details_label.configure(text=details)

    def open_logs(self):
        logfile = Path(LOGFILE)
        if not logfile.exists():
            messagebox.showinfo("Logs", "Arquivo de log não encontrado.")
            return
        try:
            if sys.platform.startswith("win"):
                os.startfile(str(logfile))
            elif sys.platform.startswith("darwin"):
                os.system(f"open '{logfile}'")
            else:
                os.system(f"xdg-open '{logfile}'")
        except Exception as e:
            messagebox.showerror("Erro", f"Não foi possível abrir logs:\n{e}")

    # ----------------- Menu / Shortcuts -----------------
    def create_menu(self):
        menubar = tkinter.Menu(self)
        filemenu = tkinter.Menu(menubar, tearoff=0)
        filemenu.add_command(label="Nova Organização (Ctrl+N)", command=self.show_creation_frame, accelerator="Ctrl+N")
        filemenu.add_command(label="Salvar (Ctrl+S)", command=self.save_organization, accelerator="Ctrl+S")
        filemenu.add_separator()
        filemenu.add_command(label="Sair", command=self.destroy)
        menubar.add_cascade(label="Arquivo", menu=filemenu)

        helpmenu = tkinter.Menu(menubar, tearoff=0)
        helpmenu.add_command(label="Abrir Logs", command=self.open_logs)
        menubar.add_cascade(label="Ajuda", menu=helpmenu)

        self.configure(menu=menubar)
        # Bind shortcuts
        self.bind_all("<Control-n>", lambda e: self.show_creation_frame())
        self.bind_all("<Control-N>", lambda e: self.show_creation_frame())
        self.bind_all("<Control-s>", lambda e: self.save_organization())
        self.bind_all("<Control-S>", lambda e: self.save_organization())

    # ----------------- Thread queue processing -----------------
    def _process_thread_queue(self):
        # placeholder if needed, already used by specific handlers
        self.after(200, self._process_thread_queue)


if __name__ == "__main__":
    app = App()

    app.mainloop()
