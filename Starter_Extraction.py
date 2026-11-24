"""P&ID step: run the P&ID through the Drawing-to-instrument-list-translator submodule's instrument list
transcription, EXACTLY as the submodule's own window does it (Instrument_List_GUI.build_args -> run_pipeline with
run mode "pids" -> transcribe_pids_drawings with the submodule's own P&ID prompt). Result: "<P&ID> - pids.json".

The submodule's functions are called as they are, never edited or patched. The options are picked in
run_pid_options() (same choices as the submodule window). To keep the job folder untouched, the P&ID is first
copied into the Working folder and the submodule runs on that copy, so its page images and the result land in
Working. (Its OCR cache stays where the submodule itself keeps it: the submodule's ocr_cache folder.)

Also holds the starter part-number table helpers (load_code_table, build_code_string, ...) used by
Starter_Selection.
"""
import json
import re
import shutil
import sys
from pathlib import Path

from GUI import WORKING_DIR, make_window, map_nested, working_path

APP_DIR = Path(__file__).resolve().parent
SUBMODULE_DIR = APP_DIR / "Submodules" / "Drawing-to-instrument-list-translator"
CODE_TABLE_FILE = APP_DIR / "Starter_Part_Number_Code.json"
PID_REFERENCE_IMAGE = SUBMODULE_DIR / "AI Reference - P&IDs - smaller.png"  # P&ID symbol legend

DEFAULT_MODEL = "qwen3-vl:4b-instruct"   # local Ollama vision model
DEFAULT_HOST = "http://localhost:11434"   # local Ollama server
DEFAULT_MAX_TOKENS = 8192                 # longest answer the model may write
DEFAULT_DPI = 300
DEFAULT_OCR_SCALE = 3.0                   # small text on drawings needs 3.0-4.0


def _squash(text) -> str:
    """Lowercase and keep only letters/digits: "P-101 A" -> "p101a". Used to match tags and names."""
    return re.sub(r"[^a-z0-9]", "", str(text).lower())


def load_code_table(path=CODE_TABLE_FILE) -> dict:
    """Read Starter_Part_Number_Code.json. `path` may be a nested list of tables (same-shape results back)."""
    if isinstance(path, (list, tuple)):
        return map_nested(load_code_table, path)
    return json.loads(Path(path).read_text(encoding="utf-8"))


def fixed_fields(table: dict) -> dict:
    """Fields with only one possible code (Customer = MY, Future Use = 0): {name: code}. Never asked of the model."""
    return {f["name"]: next(iter(f["options"])) for f in table["fields"] if len(f["options"]) == 1}


def _clean_code(field: dict, raw) -> str:
    """Normalise a code the model returned and check it is allowed; "" when it is not a valid code.
    HP codes are zero-padded (250 -> 0250)."""
    a, b = field["positions"]
    code = str(raw or "").strip().upper()
    if code.isdigit() and len(code) < b - a + 1:
        code = code.zfill(b - a + 1)
    return code if code in field["options"] else ""


def build_code_string(table: dict, chosen: dict) -> str:
    """Assemble the 20-character starter code from {field name: code}.
    Fixed fields are filled in automatically; unknown fields become '?' (one per character)."""
    fixed = fixed_fields(table)
    parts = []
    for field in table["fields"]:
        a, b = field["positions"]
        parts.append(chosen.get(field["name"]) or fixed.get(field["name"]) or "?" * (b - a + 1))
    return "".join(parts)


def _import_submodule():
    """Import the submodule's transcriber (slow: loads Docling). Its own imports expect its folder on sys.path."""
    if str(SUBMODULE_DIR) not in sys.path:
        sys.path.insert(0, str(SUBMODULE_DIR))
    import Instrument_List_Transcriber as ilt
    return ilt


def _auto_name(pdf: str, suffix: str) -> str:
    """Result file name, same as the submodule window: "<pdf name> - pids.json"."""
    stem = re.sub(r" - (controls|pids)$", "", Path(pdf).stem)
    return f"{stem} - {suffix}"


def transcribe_pid_instruments(pdf_path, model=DEFAULT_MODEL, api_host=DEFAULT_HOST, api_key="", pages=None,
                               log=print, stream=None, reference_image=PID_REFERENCE_IMAGE, dpi=None,
                               max_tokens=None, ocr_scale=DEFAULT_OCR_SCALE, use_ocr=True, ocr_tables=True,
                               ocr_pictures=False, output_path=None):
    """Run the submodule's P&ID instrument list transcription on `pdf_path`, exactly like its window's Run button
    in "P&ID drawings only" mode: the arguments are built the same way as Instrument_List_GUI.build_args and passed
    to Instrument_List_Transcriber.run_pipeline. Saves "<pdf name> - pids.json" in the Working folder (or at
    `output_path`) and returns its path.
    `max_tokens` None = the model's own limit, else 8192 (as the submodule window); `dpi` None = 300.
    `pdf_path` may be a nested list of PDFs: each is processed in turn and the paths come back in the same shape.
    `log(text)` receives progress lines and `stream(text)` the model's live output."""
    options = {k: v for k, v in locals().items() if k not in ("pdf_path", "output_path")}
    if isinstance(pdf_path, (list, tuple)):
        return map_nested(lambda one: transcribe_pid_instruments(one, **options), pdf_path)

    source = Path(pdf_path).expanduser().resolve()
    if not source.is_file():
        raise FileNotFoundError(f"P&ID PDF not found: {source}")
    # Work on a copy in Working: the submodule writes its page images beside the PDF it is given, and the job
    # folder must stay untouched. copy2 keeps the date, so the submodule's OCR cache still recognises re-runs.
    pdf = working_path(source.name)
    if pdf != source:
        shutil.copy2(source, pdf)
    log(f"P&ID copied to the Working folder: {pdf}")
    target = Path(output_path) if output_path else working_path(_auto_name(str(source), "pids.json"))

    log("Loading the instrument list transcriber (Docling)... this can take a minute the first time.")
    ilt = _import_submodule()

    # --- same as Instrument_List_GUI.resolve_endpoint: max tokens = override, else the model's own limit ---
    if not max_tokens:
        max_tokens = DEFAULT_MAX_TOKENS
        if not _claude_code().is_claude_code_model(model):
            try:
                max_tokens = int(ilt._get_model_capabilities(model, api_host, api_key, api_host == DEFAULT_HOST)
                                 .get("max_tokens") or 0) or DEFAULT_MAX_TOKENS
            except Exception:  # noqa: BLE001 - same fallback as the submodule window
                max_tokens = DEFAULT_MAX_TOKENS

    # --- same as Instrument_List_GUI.build_args for run mode "pids" ---
    args = ilt.build_arg_parser().parse_args([])
    args.run_mode = "pids"
    args.model, args.api_host, args.api_key, args.max_tokens = model, api_host, api_key, int(max_tokens)
    args.dpi = int(dpi or DEFAULT_DPI)
    for stage in ("controls", "pids", "combine"):
        setattr(args, f"{stage}_model", model)
        setattr(args, f"{stage}_api_host", api_host)
        setattr(args, f"{stage}_api_key", api_key)
        setattr(args, f"{stage}_max_tokens", int(max_tokens))
    args.pids_pdf = str(pdf)
    args.pids_pages = (pages or "").strip() or None
    args.pids_reference_image = str(reference_image) if reference_image else None
    args.pids_output = str(target)
    args.docling_enable = bool(use_ocr)
    args.docling_do_ocr = True
    args.docling_do_table_structure = bool(ocr_tables)
    args.docling_do_picture_description = bool(ocr_pictures)
    args.docling_images_scale = float(ocr_scale or DEFAULT_OCR_SCALE)

    # Same "Run settings" line the submodule window writes to its run log (without the API key).
    shown = {"mode": "pids", "pids_pdf": str(source), "pids_pages": args.pids_pages, "model": model,
             "api_host": api_host, "max_tokens": args.max_tokens, "dpi": args.dpi, "use_ocr": args.docling_enable,
             "ocr_scale": args.docling_images_scale, "ocr_tables": args.docling_do_table_structure,
             "ocr_pictures": args.docling_do_picture_description, "reference": args.pids_reference_image}
    log("Run settings: " + json.dumps({k: v for k, v in shown.items() if v not in ("", None)}, ensure_ascii=False))

    ilt.run_pipeline(args, log, stream)   # the submodule's own pipeline, unchanged
    if not target.is_file():
        raise RuntimeError(f"The transcriber finished but did not save {target}")
    log(f"P&ID instrument list: {target}")
    return target


# ------------------------------------------------------------------------------------------ model choices
CLOUD_HOSTS = {
    "Anthropic": "https://api.anthropic.com",
    "OpenAI": "https://api.openai.com",
    "Google": "https://generativelanguage.googleapis.com",
}
CLOUD_PREFIXES = ("claude", "gpt", "gemini", "o1", "o3", "o4", "chatgpt")


def _claude_code():
    """The submodule's Claude Code CLI backend (light import, unlike the transcriber)."""
    if str(SUBMODULE_DIR) not in sys.path:
        sys.path.insert(0, str(SUBMODULE_DIR))
    import Claude_Code_Backend as claude_code
    return claude_code


def _guess_provider(model: str, api_key: str) -> str:
    """Cloud provider for a model name / API key (same rules as the submodule window)."""
    lowered = model.lower()
    if lowered.startswith("claude"):
        return "Anthropic"
    if lowered.startswith(("gpt", "o1", "o3", "o4", "chatgpt")):
        return "OpenAI"
    if lowered.startswith("gemini"):
        return "Google"
    if api_key.startswith("sk-ant"):
        return "Anthropic"
    if api_key.startswith("AIza"):
        return "Google"
    return "OpenAI"


def _providers_for_key(api_key: str) -> list:
    if api_key.startswith("sk-ant"):
        return ["Anthropic"]
    if api_key.startswith("AIza"):
        return ["Google"]
    if api_key.startswith("sk-"):
        return ["OpenAI"]
    return []


def is_cloud_model(model, local_models=()) -> bool:
    """True when `model` runs on a cloud API (needs an API key). `model` may be a nested list."""
    if isinstance(model, (list, tuple)):
        return map_nested(lambda one: is_cloud_model(one, local_models), model)
    if not model or _claude_code().is_claude_code_model(model) or model in local_models:
        return False
    return model.lower().startswith(CLOUD_PREFIXES)


def model_endpoint(model: str, api_key: str = "", local_models=()) -> tuple:
    """(api_host, api_key) for a model: Claude Code CLI -> ("", ""), cloud -> (provider host, key),
    otherwise the local Ollama server."""
    if _claude_code().is_claude_code_model(model):
        return "", ""
    if is_cloud_model(model, local_models):
        return CLOUD_HOSTS[_guess_provider(model, api_key)], api_key
    return DEFAULT_HOST, ""


def list_models(api_key: str = "") -> tuple:
    """(local Ollama models, every model to offer, error text or None): local Ollama models first, then cloud
    models for the API key's provider, then the Claude Code CLI choices, like the submodule window.
    Slow parts (network) — call it from a worker thread."""
    import requests
    local, error = [], None
    try:
        response = requests.get(f"{DEFAULT_HOST}/api/tags", timeout=5)
        response.raise_for_status()
        local = sorted({m["name"] for m in response.json().get("models", []) if isinstance(m, dict) and m.get("name")},
                       key=str.lower)
    except Exception:  # noqa: BLE001
        error = "Ollama isn't reachable on this computer (is it running?). Claude Code and cloud models still work."
    cloud = []
    providers = _providers_for_key(api_key)
    if providers:
        ilt = _import_submodule()
        for provider in providers:
            cloud += [m for m in ilt._get_cloud_models(api_key, CLOUD_HOSTS[provider], provider) if m not in local]
    names = local + sorted(dict.fromkeys(cloud), key=str.lower) + list(_claude_code().available_models())
    return local, names, error


# --------------------------------------------------------------------------------------- options window
PID_SETTINGS_FILE = WORKING_DIR / "pid_settings.json"
PID_DEFAULTS = {
    "model": DEFAULT_MODEL, "api_key": "", "pages": "", "use_ocr": True, "advanced_open": False,
    "max_tokens": "", "dpi": "", "ocr_scale": str(DEFAULT_OCR_SCALE),
    "ocr_tables": True, "ocr_pictures": False, "auto_reference": True, "reference": "",
}
PAGES_PATTERN = re.compile(r"^\s*\d+(\s*-\s*\d+)?(\s*,\s*\d+(\s*-\s*\d+)?)*\s*$")


def _load_pid_settings() -> dict:
    """Last choices from the options window (Working/pid_settings.json), on top of the defaults."""
    settings = dict(PID_DEFAULTS)
    try:
        saved = json.loads(PID_SETTINGS_FILE.read_text(encoding="utf-8"))
        settings.update({k: v for k, v in saved.items() if k in PID_DEFAULTS})
    except (OSError, ValueError):
        pass
    return settings


def run_pid_options(pdf_path):
    """Options window shown before the P&ID is read: the same choices as the submodule's window for one P&ID
    stage (drawing + pages, AI model with Refresh list and API key, OCR on/off, and under Advanced: max output
    tokens, render DPI, OCR image scale, OCR tables / pictures, reference image).
    Choices are remembered in Working/pid_settings.json.
    Returns the keyword arguments for transcribe_pid_instruments() (incl. "pdf_path") when Run is pressed, or None when
    the P&ID step is skipped (Skip button or window closed).
    `pdf_path` may be a nested list of PDFs (one window each, same-shape results back)."""
    if isinstance(pdf_path, (list, tuple)):
        return map_nested(run_pid_options, pdf_path)
    import gc
    import queue
    import threading
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    settings = _load_pid_settings()
    gc.collect()  # free Tk objects from earlier windows on this (main) thread, see GUI.run_with_progress
    root = make_window("P&ID reading options", "820x780", (720, 480))
    style = ttk.Style()
    style.configure("Head.TLabel", font=("Segoe UI", 15, "bold"))
    style.configure("Section.TLabelframe.Label", font=("Segoe UI", 10, "bold"))
    style.configure("Hint.TLabel", foreground="#9a9a9a")
    style.configure("Run.TButton", font=("Segoe UI", 11, "bold"), padding=(18, 6))

    vars_ = {k: (tk.BooleanVar(value=bool(v)) if isinstance(PID_DEFAULTS[k], bool) else tk.StringVar(value=str(v)))
             for k, v in settings.items()}
    vars_["pdf"] = tk.StringVar(value=str(pdf_path or ""))
    outcome = {"result": None}
    models = {"local": [], "names": []}
    inbox: "queue.Queue" = queue.Queue()

    frame = ttk.Frame(root, padding=(20, 14))
    frame.pack(fill="both", expand=True)
    frame.columnconfigure(0, weight=1)
    ttk.Label(frame, text="P&ID reading options", style="Head.TLabel").grid(row=0, column=0, sticky="w")
    ttk.Label(frame, text="The P&ID goes through the submodule's instrument list transcriber (P&ID stage, Docling "
                          "OCR + AI model), exactly as in its own window. A copy of the P&ID is used, in the "
                          "Working folder.",
              style="Hint.TLabel", wraplength=760, justify="left").grid(row=1, column=0, sticky="w", pady=(2, 10))

    def section(row: int, title: str) -> ttk.Labelframe:
        box = ttk.Labelframe(frame, text=title, style="Section.TLabelframe", padding=(12, 8))
        box.grid(row=row, column=0, sticky="ew", pady=(0, 10))
        box.columnconfigure(1, weight=1)
        return box

    # ---- 1. drawing ----
    sec_pdf = section(2, "1. Drawing")
    ttk.Label(sec_pdf, text="P&ID PDF").grid(row=0, column=0, sticky="e", padx=(0, 8), pady=3)
    ttk.Entry(sec_pdf, textvariable=vars_["pdf"]).grid(row=0, column=1, sticky="ew", pady=3)
    ttk.Button(sec_pdf, text="Browse...", command=lambda: vars_["pdf"].set(
        filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf"), ("All files", "*.*")]) or vars_["pdf"].get())
               ).grid(row=0, column=2, padx=(8, 0))
    ttk.Label(sec_pdf, text="Pages").grid(row=1, column=0, sticky="e", padx=(0, 8), pady=3)
    ttk.Entry(sec_pdf, textvariable=vars_["pages"], width=12).grid(row=1, column=1, sticky="w", pady=3)
    ttk.Label(sec_pdf, text="Leave Pages empty for every page, or use e.g. 1,3-5.", style="Hint.TLabel").grid(
        row=2, column=1, columnspan=2, sticky="w")

    # ---- 2. model ----
    sec_model = section(3, "2. AI model")
    ttk.Label(sec_model, text="Model").grid(row=0, column=0, sticky="e", padx=(0, 8), pady=3)
    model_combo = ttk.Combobox(sec_model, textvariable=vars_["model"])
    model_combo.grid(row=0, column=1, sticky="ew", pady=3)
    refresh_btn = ttk.Button(sec_model, text="Refresh list", command=lambda: refresh_models())
    refresh_btn.grid(row=0, column=2, padx=(8, 0))
    model_note = ttk.Label(sec_model, text="", style="Hint.TLabel", wraplength=560, justify="left")
    model_note.grid(row=1, column=1, columnspan=2, sticky="w")
    key_label = ttk.Label(sec_model, text="API key")
    key_entry = ttk.Entry(sec_model, textvariable=vars_["api_key"], show="•")
    show_key = tk.BooleanVar(value=False)
    key_toggle = ttk.Checkbutton(sec_model, text="Show", variable=show_key,
                                 command=lambda: key_entry.configure(show="" if show_key.get() else "•"))

    def update_model_widgets(*_args) -> None:
        model = vars_["model"].get().strip()
        if not model:
            model_note.configure(text="Pick the vision model to use. Local models come from Ollama.")
        elif _claude_code().is_claude_code_model(model):
            model_note.configure(text="Runs through your installed Claude Code CLI (uses your Claude login; no API "
                                      "key needed)." if _claude_code().cli_path() else
                                 "The Claude Code CLI ('claude') isn't on PATH, so this model can't run.")
        elif is_cloud_model(model, models["local"]):
            model_note.configure(text=f"Cloud model ({_guess_provider(model, vars_['api_key'].get().strip())}); "
                                      "needs an API key.")
        else:
            model_note.configure(text="Runs locally on Ollama; no API key needed.")
        if is_cloud_model(model, models["local"]):
            key_label.grid(row=2, column=0, sticky="e", padx=(0, 8), pady=3)
            key_entry.grid(row=2, column=1, sticky="ew", pady=3)
            key_toggle.grid(row=2, column=2, padx=(8, 0))
        else:
            for widget in (key_label, key_entry, key_toggle):
                widget.grid_remove()

    vars_["model"].trace_add("write", update_model_widgets)
    vars_["api_key"].trace_add("write", update_model_widgets)

    def refresh_models() -> None:
        refresh_btn.configure(state="disabled")
        model_note.configure(text="Looking for models...")
        api_key = vars_["api_key"].get().strip()

        def work() -> None:  # worker thread: plain data only, never Tk objects
            try:
                inbox.put(list_models(api_key))
            except Exception as exc:  # noqa: BLE001
                inbox.put(([], [], f"Could not list models: {exc}"))

        threading.Thread(target=work, daemon=True).start()

    def pump() -> None:
        try:
            while True:
                local, names, error = inbox.get_nowait()
                models["local"], models["names"] = local, names
                model_combo.configure(values=names)
                if not vars_["model"].get().strip() and names:
                    vars_["model"].set(next((n for n in local if "vl" in n.lower() or "gemma" in n.lower()),
                                            names[0]))
                update_model_widgets()
                if error:
                    model_note.configure(text=error)
                refresh_btn.configure(state="normal")
        except queue.Empty:
            pass
        if root.winfo_exists():
            root.after(100, pump)

    # ---- 3. reading ----
    sec_read = section(4, "3. Reading")
    ttk.Checkbutton(sec_read, text="Read the drawing text with OCR first (recommended; cached, so re-runs are fast)",
                    variable=vars_["use_ocr"]).grid(row=0, column=0, columnspan=3, sticky="w")

    # ---- advanced ----
    adv_label = tk.StringVar()
    ttk.Button(frame, textvariable=adv_label, command=lambda: toggle_advanced()).grid(row=5, column=0, sticky="w")
    adv = ttk.Frame(frame)
    adv.grid(row=6, column=0, sticky="ew", pady=(6, 0))
    adv.columnconfigure(0, weight=1)
    tuning = ttk.Labelframe(adv, text="Tuning (empty = automatic)", style="Section.TLabelframe", padding=(12, 8))
    tuning.grid(row=0, column=0, sticky="ew", pady=(0, 10))
    for i, (label, key, hint) in enumerate((
        ("Max output tokens", "max_tokens", f"automatic: the model's limit (else {DEFAULT_MAX_TOKENS})"),
        ("Render DPI", "dpi", f"automatic: {DEFAULT_DPI}"),
        ("OCR image scale", "ocr_scale", "3 to 4 is plenty; memory grows with its square"),
    )):
        ttk.Label(tuning, text=label).grid(row=i, column=0, sticky="e", padx=(0, 8), pady=3)
        ttk.Entry(tuning, textvariable=vars_[key], width=12).grid(row=i, column=1, sticky="w", pady=3)
        ttk.Label(tuning, text=hint, style="Hint.TLabel").grid(row=i, column=2, sticky="w", padx=(8, 0))
    ttk.Checkbutton(tuning, text="OCR: analyse table structure", variable=vars_["ocr_tables"]).grid(
        row=4, column=1, columnspan=2, sticky="w")
    ttk.Checkbutton(tuning, text="OCR: describe pictures (slow)", variable=vars_["ocr_pictures"]).grid(
        row=5, column=1, columnspan=2, sticky="w")
    ref = ttk.Labelframe(adv, text="Reference image (notation legend sent with each page)",
                         style="Section.TLabelframe", padding=(12, 8))
    ref.grid(row=1, column=0, sticky="ew", pady=(0, 10))
    ref.columnconfigure(1, weight=1)
    ttk.Checkbutton(ref, text="Use the P&ID reference image that ships with the transcriber",
                    variable=vars_["auto_reference"]).grid(row=0, column=0, columnspan=3, sticky="w")
    ttk.Label(ref, text="P&ID image").grid(row=1, column=0, sticky="e", padx=(0, 8), pady=3)
    ttk.Entry(ref, textvariable=vars_["reference"]).grid(row=1, column=1, sticky="ew", pady=3)
    ttk.Button(ref, text="Browse...", command=lambda: vars_["reference"].set(filedialog.askopenfilename(
        filetypes=[("Images", "*.png *.jpg *.jpeg *.bmp *.tif *.tiff"), ("All files", "*.*")])
        or vars_["reference"].get())).grid(row=1, column=2, padx=(8, 0))
    ttk.Label(ref, text="A custom image here overrides the built-in one.", style="Hint.TLabel").grid(
        row=2, column=1, sticky="w")

    def apply_advanced() -> None:
        if vars_["advanced_open"].get():
            adv.grid()
            adv_label.set("▾ Advanced settings")
        else:
            adv.grid_remove()
            adv_label.set("▸ Advanced settings")

    def toggle_advanced() -> None:
        vars_["advanced_open"].set(not vars_["advanced_open"].get())
        apply_advanced()

    # ---- buttons ----
    def current() -> dict:
        return {k: v.get() for k, v in vars_.items()}

    def save_settings(cfg: dict) -> None:
        try:
            PID_SETTINGS_FILE.parent.mkdir(parents=True, exist_ok=True)
            PID_SETTINGS_FILE.write_text(json.dumps({k: cfg[k] for k in PID_DEFAULTS}, indent=2), encoding="utf-8")
        except OSError:
            pass

    def problems(cfg: dict) -> list:
        found = []
        pdf = Path(cfg["pdf"].strip()).expanduser()
        if not cfg["pdf"].strip() or not pdf.is_file():
            found.append(f"P&ID PDF not found: {cfg['pdf'].strip() or '(empty)'}")
        if cfg["pages"].strip() and not PAGES_PATTERN.match(cfg["pages"]):
            found.append("Pages must look like 1,3-5 (or be empty for every page).")
        model = cfg["model"].strip()
        if not model:
            found.append("Pick a model.")
        elif is_cloud_model(model, models["local"]) and not cfg["api_key"].strip():
            found.append("This is a cloud model, so it needs an API key.")
        for key, label in (("max_tokens", "Max output tokens"), ("dpi", "Render DPI")):
            if cfg[key].strip() and not cfg[key].strip().isdigit():
                found.append(f"{label} must be a whole number (or empty for automatic).")
        try:
            if float(cfg["ocr_scale"].strip() or DEFAULT_OCR_SCALE) <= 0:
                raise ValueError
        except ValueError:
            found.append("OCR image scale must be a positive number.")
        if cfg["reference"].strip() and not Path(cfg["reference"].strip()).expanduser().is_file():
            found.append(f"Reference image not found: {cfg['reference'].strip()}")
        return found

    def run() -> None:
        cfg = current()
        bad = problems(cfg)
        if bad:
            messagebox.showerror("Check the options", "\n".join(f"• {b}" for b in bad))
            return
        save_settings(cfg)
        model = cfg["model"].strip()
        host, key = model_endpoint(model, cfg["api_key"].strip(), models["local"])
        reference = cfg["reference"].strip() or (str(PID_REFERENCE_IMAGE) if cfg["auto_reference"] else None)
        outcome["result"] = {
            "pdf_path": str(Path(cfg["pdf"].strip()).expanduser()), "pages": cfg["pages"].strip() or None,
            "model": model, "api_host": host, "api_key": key,
            "max_tokens": int(cfg["max_tokens"]) if cfg["max_tokens"].strip() else None,
            "dpi": int(cfg["dpi"]) if cfg["dpi"].strip() else None,
            "ocr_scale": float(cfg["ocr_scale"].strip() or DEFAULT_OCR_SCALE),
            "use_ocr": cfg["use_ocr"], "ocr_tables": cfg["ocr_tables"], "ocr_pictures": cfg["ocr_pictures"],
            "reference_image": reference,
        }
        root.destroy()

    def skip() -> None:
        save_settings(current())
        root.destroy()

    # Buttons pinned to the bottom of the window (packed before the options), so they stay visible at any size.
    bar = ttk.Frame(root, padding=(20, 8, 20, 14))
    bar.pack(side="bottom", fill="x", before=frame)
    ttk.Button(bar, text="Run", style="Run.TButton", command=run).pack(side="right")
    ttk.Button(bar, text="Skip P&ID step", command=skip).pack(side="right", padx=(0, 8))
    root.bind("<Return>", lambda _e: run())
    root.protocol("WM_DELETE_WINDOW", skip)

    apply_advanced()
    update_model_widgets()
    refresh_models()
    pump()
    root.mainloop()

    # Window is gone: drop the Tk references here, on the main thread (see GUI.run_with_progress).
    del root, style, vars_, frame, model_combo, refresh_btn, model_note, key_label, key_entry, key_toggle, adv
    gc.collect()
    return outcome["result"]


