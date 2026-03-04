import webview
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path
from datetime import datetime
import re
import webbrowser

CONFIG = Path.home() / "V_emcomm.cfg"

# ------------------ Helpers ------------------

def utc_dtg() -> str:
    return datetime.utcnow().strftime("%y%m%d-%H%MZ")

def text_to_v2(s: str) -> str:
    return s.replace("\r\n", "\n").replace("\n", "[NL]")

def safe_name(s: str) -> str:
    s = (s or "").strip().replace("=", "-").replace("\n", " ").replace("\r", " ")
    s = re.sub(r"\s+", " ", s).strip()
    return s[:60] if s else "CUSTOM_FORM"

# ------------------ VarAC template V2 ops ------------------

def upsert_varac_template_v2(path: Path, name: str, subject: str, body_text: str):
    name = safe_name(name)
    subject_v2 = text_to_v2((subject or "").strip())
    body_v2 = text_to_v2((body_text or "").strip())

    value = f"{subject_v2}|{body_v2}" if subject_v2 else body_v2
    new_line = f"{name}={value}"

    if not path.exists():
        path.write_text(
            "# VarAC Templates (Forms) File\n# File version: V2\n\n" + new_line + "\n",
            encoding="utf-8"
        )
        return

    lines = path.read_text(encoding="utf-8", errors="ignore").splitlines()
    out = []
    replaced = False

    for ln in lines:
        if ln.startswith(name + "="):
            out.append(new_line)
            replaced = True
        else:
            out.append(ln)

    if not replaced:
        out.append(new_line)

    path.write_text("\n".join(out) + "\n", encoding="utf-8")

def read_varac_template_lines(path: Path):
    """
    Returns list of dicts:
      {name, state('ACTIVE'|'HIDDEN'), raw_line, clean_line}
    For VarAC V2 templates; leaves headers/comments alone.
    Hidden lines are "# NAME=..."
    """
    items = []
    if not path.exists():
        return items

    for ln in path.read_text(encoding="utf-8", errors="ignore").splitlines():
        raw = ln.rstrip()
        if not raw:
            continue

        hidden = raw.startswith("# ")
        clean = raw[2:] if hidden else raw

        # ignore non-template comments/headers
        if clean.startswith("#") or "=" not in clean:
            continue

        name = clean.split("=", 1)[0].strip()
        if not name:
            continue

        items.append({
            "name": name,
            "state": "HIDDEN" if hidden else "ACTIVE",
            "raw_line": raw,
            "clean_line": clean
        })
    return items

def write_varac_lines_preserve(path: Path, new_lines: list[str]):
    path.write_text("\n".join(new_lines) + "\n", encoding="utf-8")

# ------------------ Custom Form Parser (MCF701-style) ------------------

def parse_file_name(text: str):
    """
    Supports:
    - Header: "SUBJECT"
    - Text fields: "[TO] Box 1 - Recipient (TO):"
    - Dropdown blocks:
        "? Box 3 - Precedence:"
        "@R Routine"
        "@P Priority"
    - Multiline inferred by keywords: narrative/notes/details/remarks/summary/description
    - UTC button inferred by keywords: dtg/date/time
    """
    lines = [ln.strip("\ufeff") for ln in text.splitlines()]

    subject_default = "form_code"
    form_code = ""

    for ln in lines:
        if ln.strip() and "|" in ln and not ln.strip().startswith(("#", "!", ".")):
            a, b = ln.split("|", 1)
            subject_default = a.strip() or subject_default
            form_code = b.strip()
            break

    fields = []
    seen = set()
    i = 0
    while i < len(lines):
        ln = lines[i].strip()

        if not ln or ln.startswith(("#", "!", ".")):
            i += 1
            continue

        # Dropdown block
        if ln.startswith("?"):
            label_full = ln[1:].strip().rstrip(":")
            label = re.sub(r"^Box\s*\d+\s*-\s*", "", label_full).strip()
            code = re.sub(r"[^A-Za-z0-9]+", "_", label.upper()).strip("_") or f"DROPDOWN_{len(fields)+1}"
            if code in seen:
                n = 2
                while f"{code}_{n}" in seen:
                    n += 1
                code = f"{code}_{n}"
            seen.add(code)

            opts = []
            j = i + 1
            while j < len(lines) and lines[j].strip().startswith("@"):
                opts.append(lines[j].strip())
                j += 1

            fields.append({
                "type": "dropdown",
                "code": code,
                "label": label,
                "options": opts if opts else ["@U Unknown"],
                "multiline": False,
                "utc_button": False
            })
            i = j
            continue

        # Text field
        m = re.match(r"^\[([A-Za-z0-9]{2})\]\s*(.+):\s*$", ln)
        if m:
            code = m.group(1).upper().strip()
            label_full = m.group(2).strip()
            label = re.sub(r"^Box\s*\d+\s*-\s*", "", label_full).strip()

            if code in seen:
                n = 2
                while f"{code}{n}" in seen:
                    n += 1
                code = f"{code}{n}"
            seen.add(code)

            multiline = bool(re.search(r"(narrative|notes|details|remarks|summary|description)", label, re.I))
            utc_button = bool(re.search(r"(dtg|date|date/time|time)", label, re.I))

            fields.append({
                "type": "text",
                "code": code,
                "label": label,
                "options": None,
                "multiline": multiline,
                "utc_button": utc_button
            })
            i += 1
            continue

        i += 1

    if not form_code:
        form_code = ""

    return {"subject_default": subject_default, "form_code": form_code, "fields": fields}

# ------------------ App ------------------

class MagnetVaracControlPanel:
    def __init__(self, root):
        self.root = root
        self.root.title("V-EComm de KG5RKW")
        self.root.geometry("1100x740")
        self.root.minsize(950, 620)

        self.style = ttk.Style()
        self.theme = "dark"

        # operator identity
        self.callsign = ""
        self.winlink_templates_path = ""
        self.winlink_save_path = ""

        # paths
        self.varac_ini_path = ""
        self.bbs_folder_path = ""
        self.forms_folder_path = ""

        # custom forms
        self.form_files = []
        self.current_form = None
        self.widgets = {}     # code -> widget (Entry/Text/Combobox)
        self.drop_vars = {}   # code -> StringVar

        self.text_vars = {}   # code -> StringVar (Entry fields)

        # template manager
        self.manager_items = []

        self._build_ui()
        self._load_config()

        # Ensure CFG exists on first run
        try:
            if not CONFIG.exists():
                self._save_config(silent=True)
        except Exception:
            pass
        self._apply_theme()

        # Safe shutdown hook (after UI ready)
        self.root.after(100, lambda: self.root.protocol('WM_DELETE_WINDOW', self._on_close))

    # ---------- UI ----------
    def _build_ui(self):
        nb = ttk.Notebook(self.root)
        nb.pack(fill="both", expand=True)

        # ---------------- FORMS TAB ----------------
        self.tab_forms = ttk.Frame(nb, padding=10)
        nb.add(self.tab_forms, text="Forms")

        top = ttk.Frame(self.tab_forms)
        top.pack(fill="x")

        ttk.Label(top, text="Custom Form:").pack(side="left")
        self.form_combo = ttk.Combobox(top, state="readonly", width=55)
        self.form_combo.pack(side="left", padx=8, fill="x", expand=True)
        self.form_combo.bind("<<ComboboxSelected>>", lambda e: self._load_selected_form())

        ttk.Button(top, text="Reload Forms", command=self._reload_forms).pack(side="left")
        ttk.Button(top, text="NEW", command=self._new_clean_form).pack(side="left", padx=6)

        subj = ttk.LabelFrame(self.tab_forms, text="Subject", padding=8)
        subj.pack(fill="x", pady=8)
        self.subject_entry = ttk.Entry(subj)
        self.subject_entry.pack(fill="x")
        self.subject_entry.bind("<KeyRelease>", lambda e: self._refresh_preview())

        mid = ttk.Frame(self.tab_forms)
        mid.pack(fill="both", expand=True)

        # scrollable fields
        self.fields_container = ttk.LabelFrame(mid, text="Fill Fields", padding=4)
        self.fields_container.pack(side="left", fill="both", expand=True, padx=(0, 6))

        self.fields_canvas = tk.Canvas(self.fields_container, highlightthickness=0)
        self.fields_scroll = ttk.Scrollbar(self.fields_container, orient="vertical", command=self.fields_canvas.yview)
        self.fields_canvas.configure(yscrollcommand=self.fields_scroll.set)

        self.fields_scroll.pack(side="right", fill="y")
        self.fields_canvas.pack(side="left", fill="both", expand=True)

        self.fields_frame = ttk.Frame(self.fields_canvas)
        self.fields_canvas.create_window((0, 0), window=self.fields_frame, anchor="nw")

        self.fields_frame.bind("<Configure>", lambda e: self.fields_canvas.configure(scrollregion=self.fields_canvas.bbox("all")))
        self.fields_canvas.bind_all("<MouseWheel>", self._on_mousewheel)

        # preview
        prev = ttk.LabelFrame(mid, text="Preview (VarAC-safe text)", padding=8)
        prev.pack(side="left", fill="both", expand=True, padx=(6, 0))
        self.preview_text = tk.Text(prev, wrap="word")
        self.preview_text.pack(fill="both", expand=True)

        bottom = ttk.Frame(self.tab_forms)
        bottom.pack(fill="x", pady=8)

        ttk.Button(bottom, text="Light / Dark", command=self._toggle_theme).pack(side="left")
        ttk.Button(bottom, text="STORE BBS (.txt)", command=self._store_bbs).pack(side="left", padx=8)
        ttk.Button(bottom, text="UPDATE TEMPLATE.INI", command=self._update_template_ini).pack(side="left")

        ttk.Button(bottom, text="COPY", command=self._copy_clip).pack(side="right")

        # ---------------- SETTINGS TAB ----------------
        self.tab_settings = ttk.Frame(nb, padding=10)
        nb.add(self.tab_settings, text="Settings")

        ttk.Label(self.tab_settings, text="VarAC_templates.ini path:").pack(anchor="w")
        r1 = ttk.Frame(self.tab_settings)
        r1.pack(fill="x", pady=4)
        self.varac_entry = ttk.Entry(r1)
        self.varac_entry.pack(side="left", fill="x", expand=True)
        ttk.Button(r1, text="Browse", command=self._browse_varac_ini).pack(side="left", padx=6)

        ttk.Label(self.tab_settings, text="BBS folder (Store BBS saves .txt here):").pack(anchor="w", pady=(10, 0))
        r2 = ttk.Frame(self.tab_settings)
        r2.pack(fill="x", pady=4)
        self.bbs_entry = ttk.Entry(r2)
        self.bbs_entry.pack(side="left", fill="x", expand=True)
        ttk.Button(r2, text="Browse", command=self._browse_bbs_folder).pack(side="left", padx=6)

        ttk.Label(self.tab_settings, text="Custom forms folder (.txt forms):").pack(anchor="w", pady=(10, 0))
        r3 = ttk.Frame(self.tab_settings)
        r3.pack(fill="x", pady=4)
        self.forms_entry = ttk.Entry(r3)
        self.forms_entry.pack(side="left", fill="x", expand=True)
        ttk.Button(r3, text="Browse", command=self._browse_forms_folder).pack(side="left", padx=6)

        
        ttk.Label(self.tab_settings, text="Operator Callsign (used for filenames):").pack(anchor="w", pady=(10, 0))
        r4 = ttk.Frame(self.tab_settings)
        r4.pack(fill="x", pady=4)
        self.callsign_entry = ttk.Entry(r4)
        self.callsign_entry.pack(side="left", fill="x", expand=True)

        ttk.Button(self.tab_settings, text="Save Settings", command=self._save_config).pack(anchor="w", pady=12)
        
        ttk.Label(self.tab_settings, text="Winlink Templates Folder (.html):").pack(anchor="w", pady=(10,0))
        r5 = ttk.Frame(self.tab_settings)
        r5.pack(fill="x", pady=4)
        self.winlink_tpl_entry = ttk.Entry(r5)
        self.winlink_tpl_entry.pack(side="left", fill="x", expand=True)
        ttk.Button(r5, text="Browse", command=self._browse_winlink_templates).pack(side="left", padx=6)

        self.status_label = ttk.Label(self.tab_settings, text="")
        self.status_label.pack(anchor="w", pady=6)

        
        # ---------------- WINLINK TAB ----------------
        self.tab_winlink = ttk.Frame(nb, padding=10)
        nb.add(self.tab_winlink, text="Winlink")

        topw = ttk.Frame(self.tab_winlink)
        topw.pack(fill="x")

        ttk.Label(topw, text="Winlink HTML:").pack(side="left")
        self.winlink_combo = ttk.Combobox(topw, state="readonly", width=50)
        self.winlink_combo.pack(side="left", padx=6, fill="x", expand=True)

        ttk.Button(topw, text="Open Selected HTML", command=self._load_winlink_html).pack(side="left")
        ttk.Button(topw, text="STORE BBS", command=self._store_bbs).pack(side="left", padx=6)

        viewer_frame = ttk.LabelFrame(self.tab_winlink, text="Winlink Form", padding=8)
        viewer_frame.pack(fill="both", expand=True, pady=10)

        self.winlink_status = ttk.Label(viewer_frame, text="No form loaded.")
        self.winlink_status.pack(anchor="w")

        ttk.Button(self.tab_winlink, text="Open in Browser (use form Save/Load)", command=self._open_winlink_in_browser).pack(anchor="e")

        # ---------------- TEMPLATE MANAGER TAB ----------------
        self.tab_manager = ttk.Frame(nb, padding=10)
        nb.add(self.tab_manager, text="Template Manager")

        self._build_manager_ui()

    def _build_manager_ui(self):
        top = ttk.Frame(self.tab_manager)
        top.pack(fill="both", expand=True)

        # left list
        left = ttk.Frame(top)
        left.pack(side="left", fill="y", padx=(0, 8))

        ttk.Label(left, text="Templates in VarAC_templates.ini").pack(anchor="w")
        self.tpl_list = tk.Listbox(left, height=28)
        self.tpl_list.pack(fill="y", expand=True)
        self.tpl_list.bind("<<ListboxSelect>>", lambda e: self._manager_preview_selected())

        ttk.Button(left, text="Reload List", command=self._manager_load_list).pack(fill="x", pady=6)

        # right preview/actions
        right = ttk.Frame(top)
        right.pack(side="left", fill="both", expand=True)

        ttk.Label(right, text="Preview").pack(anchor="w")
        self.tpl_preview = tk.Text(right, wrap="word")
        self.tpl_preview.pack(fill="both", expand=True)

        actions = ttk.Frame(right)
        actions.pack(fill="x", pady=8)

        ttk.Button(actions, text="HIDE", command=lambda: self._manager_toggle(hidden=True)).pack(side="left")
        ttk.Button(actions, text="RESTORE", command=lambda: self._manager_toggle(hidden=False)).pack(side="left", padx=8)
        ttk.Button(actions, text="DELETE", command=self._manager_delete).pack(side="left")

        ttk.Label(right, text="Note: Reopen VarAC Templates/VMail window after changes.").pack(anchor="w", pady=(10, 0))

    # ---------- Mousewheel ----------
    def _on_mousewheel(self, event):
        try:
            self.fields_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        except Exception:
            pass

    # ---------- Theme ----------
    def _apply_theme(self):
        dark = {"bg":"#1e1e1e","fg":"#e6e6e6","field":"#2b2b2b","panel":"#232323"}
        light = {"bg":"#f3f3f3","fg":"#111111","field":"#ffffff","panel":"#ffffff"}
        pal = dark if self.theme == "dark" else light

        self.root.configure(bg=pal["bg"])
        self.style.theme_use("clam")
        self.style.configure(".", background=pal["bg"], foreground=pal["fg"])
        self.style.configure("TFrame", background=pal["bg"])
        self.style.configure("TLabelframe", background=pal["bg"], foreground=pal["fg"])
        self.style.configure("TLabelframe.Label", background=pal["bg"], foreground=pal["fg"])
        self.style.configure("TLabel", background=pal["bg"], foreground=pal["fg"])
        self.style.configure("TButton", background=pal["panel"], foreground=pal["fg"])
        self.style.configure("TEntry", fieldbackground=pal["field"], foreground=pal["fg"])
        self.style.configure("TCombobox", fieldbackground=pal["field"], foreground=pal["fg"])
        # ---- Notebook (Tabs) Styling ----
        self.style.configure(
            "TNotebook",
            background=pal["bg"],
            borderwidth=0
        )

        self.style.configure(
            "TNotebook.Tab",
            background="#2b2b2b",
            foreground="#cccccc",
            padding=(12, 6),
            relief="flat"
        )

        self.style.map(
            "TNotebook.Tab",
            background=[
                ("selected", "#3f3f3f"),
                ("active", "#353535")
            ],
            foreground=[
                ("selected", "#ffffff"),
                ("active", "#ffffff")
            ]
        )


        self.preview_text.configure(bg=pal["field"], fg=pal["fg"], insertbackground=pal["fg"])
        self.tpl_preview.configure(bg=pal["field"], fg=pal["fg"], insertbackground=pal["fg"])
        self.fields_canvas.configure(bg=pal["bg"], highlightbackground=pal["bg"])

        # any Text widgets in form
        for child in self.fields_frame.winfo_children():
            if isinstance(child, tk.Text):
                child.configure(bg=pal["field"], fg=pal["fg"], insertbackground=pal["fg"])

    def _toggle_theme(self):
        self.theme = "light" if self.theme == "dark" else "dark"
        self._apply_theme()
        self._save_config(silent=True)

    # ---------- Config ----------
    def _load_config(self):
        if not CONFIG.exists():
            return
        for ln in CONFIG.read_text(encoding="utf-8", errors="ignore").splitlines():
            if ln.startswith("VARAC="):
                self.varac_ini_path = ln.split("=", 1)[1].strip()
            elif ln.startswith("BBS="):
                self.bbs_folder_path = ln.split("=", 1)[1].strip()
            elif ln.startswith("FORMS="):
                self.forms_folder_path = ln.split("=", 1)[1].strip()
            elif ln.startswith("CALLSIGN="):
                self.callsign = ln.split("=", 1)[1].strip()
            elif ln.startswith("THEME="):
                self.theme = (ln.split("=", 1)[1].strip() or "dark")
            elif ln.startswith("WINLINK_TPL="):
                self.winlink_templates_path = ln.split("=", 1)[1].strip()
            elif ln.startswith("WINLINK_SAVE="):
                self.winlink_save_path = ln.split("=", 1)[1].strip()

        if self.varac_ini_path:
            self.varac_entry.insert(0, self.varac_ini_path)
        if self.bbs_folder_path:
            self.bbs_entry.insert(0, self.bbs_folder_path)
        if self.forms_folder_path:
            self.forms_entry.insert(0, self.forms_folder_path)
        if self.callsign:
            self.callsign_entry.insert(0, self.callsign)

        if getattr(self, "winlink_templates_path", ""):
            self.winlink_tpl_entry.insert(0, self.winlink_templates_path)
            self._reload_winlink_templates()

        if getattr(self, "winlink_save_path", ""):
            getattr(self, 'winlink_save_entry', None).insert(0, self.winlink_save_path)

        self._reload_forms()
        self._manager_load_list()

    def _save_config(self, silent=False):
        self.varac_ini_path = self.varac_entry.get().strip()
        self.bbs_folder_path = self.bbs_entry.get().strip()
        self.forms_folder_path = self.forms_entry.get().strip()
        self.callsign = self.callsign_entry.get().strip().upper()
        self.winlink_templates_path = self.winlink_tpl_entry.get().strip()
        self.winlink_save_path = getattr(self, "winlink_save_path", "")

        lines = [
            f"VARAC={self.varac_ini_path}",
            f"CALLSIGN={self.callsign}",
            f"BBS={self.bbs_folder_path}",
            f"FORMS={self.forms_folder_path}",
            f"THEME={self.theme}",
            f"WINLINK_TPL={self.winlink_templates_path}",
            f"WINLINK_SAVE={self.winlink_save_path}",
        ]
        try:
            CONFIG.write_text("\n".join(lines) + "\n", encoding="utf-8")
        except Exception as e:
            if not silent:
                messagebox.showerror("Config write failed", f"Could not write config to:\n{CONFIG}\n\nError: {e}")
            return
        if not silent:
            self.status_label.config(text="Settings saved.")
            self._reload_forms()
            self._reload_winlink_templates()
            self._manager_load_list()

    # ---------- Browsers ----------
    def _browse_varac_ini(self):
        f = filedialog.askopenfilename(title="Select VarAC_templates.ini", filetypes=[("INI Files", "*.ini"), ("All files", "*.*")])
        if f:
            self.varac_entry.delete(0, "end")
            self.varac_entry.insert(0, f)
            self._manager_load_list()

    def _browse_bbs_folder(self):
        d = filedialog.askdirectory(title="Select BBS folder")
        if d:
            self.bbs_entry.delete(0, "end")
            self.bbs_entry.insert(0, d)

    def _browse_forms_folder(self):
        d = filedialog.askdirectory(title="Select forms folder")
        if d:
            self.forms_entry.delete(0, "end")
            self.forms_entry.insert(0, d)
            self._reload_forms()

    # ---------- Forms ----------
    def _reload_forms(self):
        folder = Path(self.forms_entry.get().strip()) if self.forms_entry.get().strip() else None
        self.form_files = []
        if folder and folder.exists():
            self.form_files = sorted([p for p in folder.glob("*.txt") if p.is_file()])

        names = [p.name for p in self.form_files]
        self.form_combo["values"] = names
        if names:
            self.form_combo.current(0)
            self._load_selected_form()

    def _load_selected_form(self):
        if not self.form_files:
            return
        fname = self.form_combo.get()
        fp = next((p for p in self.form_files if p.name == fname), None)
        if not fp:
            return

        txt = fp.read_text(encoding="utf-8", errors="ignore")
        parsed = parse_file_name(txt)
        # Use filename as default template name if none provided
        if not parsed.get("form_code") or parsed["form_code"] == "CUSTOM_FORM":
            parsed["form_code"] = safe_name(fp.stem)
        self.current_form = parsed

        self.subject_entry.delete(0, "end")
        self.subject_entry.insert(0, parsed["subject_default"])

        self._build_fields(parsed)
        self._refresh_preview()
        self._apply_theme()

    def _build_fields(self, parsed):
        for w in self.fields_frame.winfo_children():
            w.destroy()
        self.widgets.clear()
        self.drop_vars.clear()
        self.text_vars.clear()

        title = ttk.Label(self.fields_frame, text=f"{parsed['subject_default']}  ({parsed['form_code']})", font=("Segoe UI", 11, "bold"))
        title.pack(anchor="w", pady=(6, 10))

        for field in parsed["fields"]:
            row = ttk.Frame(self.fields_frame)
            row.pack(fill="x", pady=4)

            ttk.Label(row, text=f"{field['code']} - {field['label']}").pack(side="left", anchor="w")

            if field.get("utc_button"):
                ttk.Button(row, text="UTC", command=lambda c=field["code"]: self._set_utc_for_code(c)).pack(side="right")

            if field["type"] == "dropdown":
                var = tk.StringVar(value=field["options"][0] if field["options"] else "")
                cmb = ttk.Combobox(self.fields_frame, state="readonly", values=field["options"], textvariable=var)
                cmb.pack(fill="x")
                cmb.configure(postcommand=lambda c=cmb: c.tk.call("tk::PlaceWindow", c, "pointer"))
                var.trace_add("write", lambda *a: self._refresh_preview())
                self.drop_vars[field["code"]] = var
                self.widgets[field["code"]] = cmb
            else:
                if field.get("multiline"):
                    t = tk.Text(self.fields_frame, height=4, wrap="word")
                    t.pack(fill="x")
                    t.bind("<KeyRelease>", lambda e: self._refresh_preview())
                    self.widgets[field["code"]] = t
                else:
                    v = tk.StringVar()
                    ent = ttk.Entry(self.fields_frame, textvariable=v)
                    ent.pack(fill="x")
                    # Keep a reference to the StringVar so it can't be garbage-collected.
                    self.text_vars[field["code"]] = v
                    # Update preview on typing/paste
                    ent.bind("<KeyRelease>", lambda e: self._refresh_preview())
                    v.trace_add("write", lambda *a: self._refresh_preview())
                    self.widgets[field["code"]] = ent

        self.fields_canvas.configure(scrollregion=self.fields_canvas.bbox("all"))

    def _set_utc_for_code(self, code: str):
        w = self.widgets.get(code)
        val = utc_dtg()
        if isinstance(w, tk.Text):
            w.delete("1.0", "end")
            w.insert("1.0", val)
        elif isinstance(w, ttk.Combobox):
            pass
        else:
            try:
                w.delete(0, "end")
                w.insert(0, val)
            except Exception:
                pass
        self._refresh_preview()

    def _new_clean_form(self):
        if not self.current_form:
            return
        self.subject_entry.delete(0, "end")
        self.subject_entry.insert(0, self.current_form.get("subject_default", ""))

        for field in self.current_form["fields"]:
            code = field["code"]
            if field["type"] == "dropdown":
                if field["options"]:
                    self.drop_vars[code].set(field["options"][0])
            else:
                w = self.widgets.get(code)
                if isinstance(w, tk.Text):
                    w.delete("1.0", "end")
                else:
                    try:
                        w.delete(0, "end")
                    except Exception:
                        pass
        self._refresh_preview()

    def _collect_values(self):
        vals = {}
        if not self.current_form:
            return vals
        for field in self.current_form["fields"]:
            code = field["code"]
            if field["type"] == "dropdown":
                vals[code] = self.drop_vars.get(code).get().strip()
            else:
                w = self.widgets.get(code)
                if isinstance(w, tk.Text):
                    vals[code] = w.get("1.0", "end").strip()
                else:
                    try:
                        vals[code] = w.get().strip()
                    except Exception:
                        vals[code] = ""
        return vals

    def _refresh_preview(self):
        if not self.current_form:
            return
        subject = self.subject_entry.get().strip()
        vals = self._collect_values()

        lines = []
        if subject:
            lines.append(f"SUBJECT: {subject}")
            lines.append("")

        for field in self.current_form["fields"]:
            code = field["code"]
            label = field["label"]
            v = (vals.get(code, "") or "").strip()

            if field["type"] == "text" and field.get("multiline"):
                lines.append(f"{code} {label}:")
                if v:
                    lines.append(v)
                lines.append("")
            else:
                lines.append(f"{code} {label}: {v}".rstrip())

        preview = "\n".join(lines).rstrip() + "\n"
        self.preview_text.delete("1.0", "end")
        self.preview_text.insert("1.0", preview)

    # ---------- Actions ----------
    def _copy_clip(self):
        txt = self.preview_text.get("1.0", "end").strip()
        if not txt:
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(txt)
        self.root.update()
        messagebox.showinfo("Copied", "Copied to clipboard.")

    def _store_bbs(self):
        folder = Path(self.bbs_entry.get().strip()) if self.bbs_entry.get().strip() else None
        if not folder or not folder.exists():
            messagebox.showerror("BBS Folder", "Set a valid BBS folder in Settings.")
            return
        txt = self.preview_text.get("1.0", "end").strip()
        if not txt:
            return
        form_code = self.current_form.get("form_code", "FORM") if self.current_form else "FORM"

        # --- Auto filename tagging ---
        callsign = "STATION"
        try:
            values = self._collect_values()
            for key, val in values.items():
                if re.search(r"(call|from|operator|station)", key, re.I) and val.strip():
                    callsign = safe_name(val.strip().split()[0].upper())
                    break
        except Exception:
            pass

        cs = safe_name(self.callsign) if self.callsign else "STATION"
        filename = f"{safe_name(form_code)}_{cs}_{utc_dtg()}.txt"
        out = folder / filename
        out.write_text(txt, encoding="utf-8")
        messagebox.showinfo("Stored", f"Saved:\n{out}")

    def _update_template_ini(self):
        varac = Path(self.varac_entry.get().strip()) if self.varac_entry.get().strip() else None
        if not varac:
            messagebox.showerror("VarAC Path", "Set VarAC_templates.ini path in Settings.")
            return
        if not self.current_form:
            messagebox.showerror("No Form", "Load a custom form first.")
            return
        subject = self.subject_entry.get().strip()
        body = self.preview_text.get("1.0", "end").strip()
        if not body:
            return
        name = self.current_form.get("form_code", "CUSTOM_FORM")
        upsert_varac_template_v2(varac, name=name, subject=subject, body_text=body)
        messagebox.showinfo("Updated", f"Updated VarAC template:\n{safe_name(name)}\n\nReopen VarAC Templates/VMail window to reload.")
        self._manager_load_list()

    # ---------- Template Manager ----------
    def _manager_load_list(self):
        self.tpl_list.delete(0, "end")
        self.tpl_preview.delete("1.0", "end")
        self.manager_items = []

        varac = Path(self.varac_entry.get().strip()) if self.varac_entry.get().strip() else None
        if not varac or not varac.exists():
            self.tpl_list.insert("end", "(Set VarAC_templates.ini path in Settings)")
            return

        self.manager_items = read_varac_template_lines(varac)
        if not self.manager_items:
            self.tpl_list.insert("end", "(No templates found)")
            return

        for item in self.manager_items:
            self.tpl_list.insert("end", f"{item['name']} [{item['state']}]")

    def _manager_preview_selected(self):
        sel = self.tpl_list.curselection()
        if not sel or not self.manager_items:
            return
        idx = sel[0]
        if idx >= len(self.manager_items):
            return
        raw = self.manager_items[idx]["raw_line"]
        self.tpl_preview.delete("1.0", "end")
        self.tpl_preview.insert("1.0", raw)

    def _manager_toggle(self, hidden: bool):
        sel = self.tpl_list.curselection()
        if not sel or not self.manager_items:
            return
        idx = sel[0]
        if idx >= len(self.manager_items):
            return
        name = self.manager_items[idx]["name"]

        varac = Path(self.varac_entry.get().strip())
        lines = varac.read_text(encoding="utf-8", errors="ignore").splitlines()
        new_lines = []
        for ln in lines:
            clean = ln[2:] if ln.startswith("# ") else ln
            if clean.startswith(name + "="):
                new_lines.append(("# " + clean) if hidden else clean)
            else:
                new_lines.append(ln)

        write_varac_lines_preserve(varac, new_lines)
        self._manager_load_list()

    def _manager_delete(self):
        sel = self.tpl_list.curselection()
        if not sel or not self.manager_items:
            return
        idx = sel[0]
        if idx >= len(self.manager_items):
            return
        name = self.manager_items[idx]["name"]

        if not messagebox.askyesno("Confirm Delete", f"Permanently delete '{name}'?"):
            return

        varac = Path(self.varac_entry.get().strip())
        lines = varac.read_text(encoding="utf-8", errors="ignore").splitlines()
        new_lines = []
        for ln in lines:
            clean = ln[2:] if ln.startswith("# ") else ln
            if clean.startswith(name + "="):
                continue
            new_lines.append(ln)

        write_varac_lines_preserve(varac, new_lines)
        self._manager_load_list()



    def _on_close(self):
        try:
            self._save_config(silent=True)
        except Exception:
            pass
        self.root.destroy()

    # ---------- Winlink Browsers ----------

    def _browse_winlink_templates(self):
        d = filedialog.askdirectory(title="Select Winlink Templates Folder")
        if d:
            self.winlink_tpl_entry.delete(0,"end")
            self.winlink_tpl_entry.insert(0,d)
            self.winlink_templates_path = d
            self._reload_winlink_templates()
            self._save_config(silent=True)

    def _browse_winlink_save(self):
        d = filedialog.askdirectory(title="Select Winlink Save Folder")
        if d:
            getattr(self, 'winlink_save_entry', None).delete(0,"end")
            getattr(self, 'winlink_save_entry', None).insert(0,d)
            self.winlink_save_path = d
            self._save_config(silent=True)

    # ---------- Winlink Logic ----------

    def _reload_winlink_templates(self):
        self.winlink_templates_path = self.winlink_tpl_entry.get().strip() or self.winlink_templates_path
        folder = Path(self.winlink_templates_path) if self.winlink_templates_path else None
        files = []
        if folder and folder.exists():
            files = sorted([p.name for p in folder.glob("*.htm*")])
        self.winlink_combo["values"] = files
        if files:
            self.winlink_combo.current(0)

    
    def _load_winlink_html(self):
        if not self.winlink_templates_path:
            messagebox.showerror("Winlink", "Set Winlink Templates Folder in Settings.")
            return

        name = self.winlink_combo.get()
        if not name:
            messagebox.showerror("Winlink", "Select a Winlink HTML form from the dropdown.")
            return

        html_path = Path(self.winlink_templates_path) / name
        if not html_path.exists():
            messagebox.showerror("Winlink", f"Template not found:\n{html_path}")
            return

        self.winlink_current_file = html_path
        self.winlink_status.config(text=f"Selected: {name}")

        # Open in the user's default browser. This preserves native Winlink form behavior
        # (the form's own Save/Load buttons, file dialogs, downloads, etc.).
        self._open_winlink_in_browser()


    
    def _open_winlink_in_browser(self):
        if not hasattr(self, "winlink_current_file") or not self.winlink_current_file:
            messagebox.showerror("Winlink", "Open a Winlink HTML form first.")
            return
        try:
            url = Path(self.winlink_current_file).resolve().as_uri()
            webbrowser.open(url)
            self.winlink_status.config(text=f"Opened in browser: {Path(self.winlink_current_file).name}")
        except Exception as e:
            messagebox.showerror("Winlink", f"Failed to open browser:\n{e}")

    def _save_winlink_form(self):
        if not hasattr(self, "winlink_current_file"):
            messagebox.showerror("Winlink", "Load a Winlink form first.")
            return

        # Read current Winlink HTML
        data = Path(self.winlink_current_file).read_text(
            encoding="utf-8",
            errors="ignore"
        )

        if not data:
            return

        save_dir = Path(self.winlink_save_path) if self.winlink_save_path else None
        if not save_dir or not save_dir.exists():
            messagebox.showerror("Winlink Save", "Set Winlink save folder in Settings.")
            return

        # ---------------------------------------------------
        # Use Winlink Form Name as Filename
        # ---------------------------------------------------
        form_name = Path(self.winlink_current_file).stem

        # operator callsign tagging
        cs = safe_name(self.callsign) if self.callsign else "STATION"

        filename = f"{safe_name(form_name)}_{cs}_{utc_dtg()}.txt"

        out = save_dir / filename
        out.write_text(data, encoding="utf-8")

        messagebox.showinfo("Winlink Saved", f"Saved:\n{out}")

        # Optional VarAC BBS copy
        if messagebox.askyesno("Store VarAC BBS", "Do you want to store VarAC BBS copy?"):
            self.preview_text.delete("1.0", "end")
            self.preview_text.insert("1.0", data)
            self._store_bbs()


# ------------------ Run ------------------

if __name__ == "__main__":
    root = tk.Tk()
    app = MagnetVaracControlPanel(root)
    root.mainloop()