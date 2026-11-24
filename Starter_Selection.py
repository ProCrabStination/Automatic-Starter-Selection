"""Select the motor starter part number (Starter_Part_Number_Code.json) from the filled turn-in values JSON.

extracted_items()        every item in the "<turn-in> - turn-in values.json" data (fields and table rows), with
                         the sheet and cell it came from and which part-number field it is used for.
resolve_inputs()         the values the rules need, taken from those items (first filled source wins).
select_codes()           applies one rule per part-number field to those inputs -> {field: (code, reason)}.
check_conditions()       warnings for site conditions that may need a derated starter (altitude, ambient).
run_starter_selection()  window listing EVERY extracted item (editable) and every part-number field (editable),
                         with the part number updating live. Saves "<turn-in> - starter selection.json" in Working.
"""
import datetime
import gc
import json
import re
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk

from openpyxl.utils import get_column_letter, range_boundaries

from GUI import make_window, map_nested, working_path
from Starter_Extraction import CODE_TABLE_FILE, _squash, build_code_string, load_code_table

APP_DIR = Path(__file__).resolve().parent
SELECTION_FILE_SUFFIX = " - starter selection"   # "<turn-in name> - starter selection.json"
NOT_SELECTED = "(not selected)"

# The turn-in values the selection rules use. "sources" are (sheet, field name) in the turn-in values JSON; the
# first one with a value is used. "used_for" is shown beside that item in the window.
INPUTS = [
    {"key": "package_type", "label": "Package Type", "used_for": "Compressor",
     "sources": [("Distribution Page", "Package Type")]},
    {"key": "screw_qty", "label": "Screw Compressor Quantity", "used_for": "Compressor",
     "sources": [("Screw Comp", "Quantity of Units")]},
    {"key": "recip_qty", "label": "Recip Compressor Quantity", "used_for": "Compressor",
     "sources": [("RU Comp", "Quantity of Units")]},
    {"key": "starting_method", "label": "Starter Starting Method", "used_for": "Drive Type",
     "sources": [("Screw Comp", "Motor Starter Starting Method"), ("Screw Comp", "Motor Startup Type")]},
    {"key": "service_factor", "label": "Motor Service Factor", "used_for": "Drive Type",
     "sources": [("Screw Comp", "Compressor Motor Service Factor")]},
    {"key": "motor_hp", "label": "Compressor Motor HP", "used_for": "HP",
     "sources": [("Screw Comp", "Compressor Motor Horse Power")]},
    {"key": "line_voltage", "label": "Starter / Motor Voltage", "used_for": "Line Voltage, Motor Speed (Hz)",
     "sources": [("Screw Comp", "Compressor Motor Starter Voltage"), ("Screw Comp", "Voltage - Compressor Motor"),
                 ("RU Comp", "Voltage - Compressor Motor")]},
    {"key": "control_voltage", "label": "Control Panel Voltage", "used_for": "Control Voltage",
     "sources": [("Screw Comp", "Voltage - Control Panel"), ("RU Comp", "Voltage - Control Panel")]},
    {"key": "space_heaters", "label": "Motor Space Heaters", "used_for": "Heater",
     "sources": [("Screw Comp", "Motor Space Heaters")]},
    {"key": "motor_speed", "label": "Compressor Motor Speed (RPM)", "used_for": "Motor Speed",
     "sources": [("Screw Comp", "Compressor Motor Speed")]},
    {"key": "starter_enclosure", "label": "Starter Enclosure", "used_for": "UL Rating",
     "sources": [("Screw Comp", "Motor Starter Enclosure")]},
    {"key": "altitude", "label": "Altitude", "used_for": "Warning (derating)",
     "sources": [("Order Information", "Altitude")]},
    {"key": "altitude_units", "label": "Altitude Units", "used_for": "Warning (derating)",
     "sources": [("Order Information", "Altitude Units")]},
    {"key": "ambient", "label": "Ambient Temperature", "used_for": "Warning (derating)",
     "sources": [("Order Information", "Ambient Temperature")]},
    {"key": "ambient_units", "label": "Ambient Temperature Units", "used_for": "Warning (derating)",
     "sources": [("Order Information", "Ambient Temperature units")]},
]


# ------------------------------------------------------------------------------------------------ inputs
def _item_key(sheet, name) -> str:
    """Key of one extracted item: sheet + field name, ignoring case/spaces/punctuation."""
    return f"{_squash(sheet)}|{_squash(name)}"


def _text(value) -> str:
    """Cell value as the text shown in an entry box: None -> "", 350 -> "350", 1.15 -> "1.15"."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value)


def _load_values(values) -> dict:
    """The turn-in values JSON as a dict (pass the dict itself or its path)."""
    return values if isinstance(values, dict) else json.loads(Path(values).read_text(encoding="utf-8"))


def extracted_items(values) -> list:
    """Every item in the turn-in values JSON, in sheet order:
    [{"key", "sheet", "name", "cell", "value" (text), "used_for"}].
    Fields come as they are; each table row becomes one item named after its label columns
    ("Absorbed Power [HP]") with the last column as its value. Items on skipped sheets are listed empty.
    `values` (dict or path) may be a nested list (same-shape results back)."""
    if isinstance(values, (list, tuple)):
        return map_nested(extracted_items, values)
    values = _load_values(values)
    used: dict = {}
    for spec in INPUTS:
        for sheet, name in spec["sources"]:
            used.setdefault(_item_key(sheet, name), []).append(spec["used_for"])
    items = []
    for block in values.get("sheets", []):
        sheet = block["sheet"]
        for field in block.get("fields", []):
            key = _item_key(sheet, field["name"])
            items.append({"key": key, "sheet": sheet, "name": field["name"], "cell": field.get("cell", ""),
                          "value": _text(field.get("value")), "used_for": ", ".join(dict.fromkeys(used.get(key, [])))})
        table = block.get("table") or {}
        if table.get("rows"):
            last_col = range_boundaries(table["range"])[2]
            for row in table["rows"]:
                columns = [c for c in row if c != "Excel Row"]
                extra = [_text(row[c]) for c in columns[1:-1] if _text(row[c]) not in ("", "-")]
                name = _text(row[columns[0]]) + (f" [{', '.join(extra)}]" if extra else "")
                items.append({"key": _item_key(sheet, f"row {row['Excel Row']}"), "sheet": sheet, "name": name,
                              "cell": f"{get_column_letter(last_col)}{row['Excel Row']}",
                              "value": _text(row[columns[-1]]) if len(columns) > 1 else "", "used_for": ""})
    return items


def pid_items(pid_result) -> list:
    """Every instrument in the submodule's P&ID instrument list ("<P&ID> - pids.json", dict or path), in the same
    shape as extracted_items(): one group per page ("P&ID - page N"), one item per instrument (name = the tag plus
    any other fields the transcriber returned, such as "P&ID Notes"; value = Responsibility), then the page's
    carry-forward notes. Returns [] when there is no P&ID result.
    `pid_result` may be a nested list (same-shape results back)."""
    if isinstance(pid_result, (list, tuple)):
        return map_nested(pid_items, pid_result)
    if not pid_result:
        return []
    pid = pid_result if isinstance(pid_result, dict) else json.loads(Path(pid_result).read_text(encoding="utf-8"))
    items = []
    for page in pid.get("pages", []):
        number = page.get("page")
        sheet = f"P&ID - page {number}"
        for index, instrument in enumerate(page.get("instruments") or [], start=1):
            if not isinstance(instrument, dict):
                continue
            tag = _text(instrument.get("name")) or "[unclear]"
            extra = "; ".join(f"{k}: {_text(v)}" for k, v in instrument.items()
                              if k not in ("name", "Responsibility") and _text(v))
            items.append({"key": _item_key(sheet, f"{index} {tag}"), "sheet": sheet,
                          "name": tag + (f"  ({extra})" if extra else ""), "cell": f"page {number}",
                          "value": _text(instrument.get("Responsibility")), "used_for": "", "instrument": True})
        notes = page.get("carry_forward") or []
        for index, note in enumerate([notes] if isinstance(notes, str) else notes, start=1):
            items.append({"key": _item_key(sheet, f"note {index}"), "sheet": sheet, "name": "Note",
                          "cell": f"page {number}", "value": _text(note), "used_for": ""})
    return items


def resolve_inputs(current) -> dict:
    """{INPUTS key: text} from the current item values ({item key: text}): for each input, the first of its
    sources that is filled. `current` may be a nested list (same-shape results back)."""
    if isinstance(current, (list, tuple)):
        return map_nested(resolve_inputs, current)
    inputs = {}
    for spec in INPUTS:
        texts = [current.get(_item_key(sheet, name), "") for sheet, name in spec["sources"]]
        inputs[spec["key"]] = next((t for t in texts if str(t).strip()), "")
    return inputs


# ----------------------------------------------------------------------------------------------- parsing
def _number(text):
    """First number in the text ("350 HP" -> 350.0, "1,200" -> 1200.0); None if there is none.
    `text` may be a nested list (same-shape results back)."""
    if isinstance(text, (list, tuple)):
        return map_nested(_number, text)
    match = re.search(r"-?\d+(?:\.\d+)?", str(text or "").replace(",", ""))
    return float(match.group()) if match else None


def _numbers(text) -> list:
    """Every number in the text: "70.0/0.00" -> [70.0, 0.0]."""
    return [float(n) for n in re.findall(r"-?\d+(?:\.\d+)?", str(text or "").replace(",", ""))]


def _volts_hz(text):
    """(volts, hz) from text like "460v/3p/60Hz" or "460V 3PH 60HZ"; None for a part not found.
    `text` may be a nested list (same-shape results back)."""
    if isinstance(text, (list, tuple)):
        return map_nested(_volts_hz, text)
    text = str(text or "")
    volts = re.search(r"(\d{2,4})\s*v", text, re.I) or re.search(r"\d{3}", text)
    hz = re.search(r"(\d{2})\s*hz", text, re.I)
    return (float(volts.group(1) if volts.lastindex else volts.group()) if volts else None,
            float(hz.group(1)) if hz else None)


def _yes_no(text):
    """True for yes / y / true / included, False for no / n / none / false / not included, None otherwise.
    `text` may be a nested list (same-shape results back)."""
    if isinstance(text, (list, tuple)):
        return map_nested(_yes_no, text)
    word = _squash(text)
    if word in {"yes", "y", "true", "included", "required", "1"}:
        return True
    if word in {"no", "n", "none", "false", "notincluded", "notrequired", "0"}:
        return False
    return None


# ------------------------------------------------------------------------------------------------- rules
# One rule per part-number field: rule(inputs, chosen, field) -> (code or "", reason). `inputs` is {key: text};
# `chosen` holds the codes already decided for earlier fields (including ones set by hand).
def _rule_compressor(i, chosen, field):
    package = _squash(i.get("package_type"))
    if package.startswith("screw"):
        return "S", f"Package Type is '{i['package_type']}'"
    if package.startswith("recip"):
        return "R", f"Package Type is '{i['package_type']}'"
    if (_number(i.get("screw_qty")) or 0) > 0:
        return "S", "Screw Comp quantity is above 0"
    if (_number(i.get("recip_qty")) or 0) > 0:
        return "R", "RU Comp quantity is above 0"
    return "", "Package Type not recognised: choose"


def _rule_drive_type(i, chosen, field):
    method = str(i.get("starting_method") or "").lower()
    if not method.strip():
        return "", "No starting method in the turn-in: choose"
    sf = _number(i.get("service_factor"))
    high_sf = sf is not None and sf >= 1.15
    vfd = any(word in method for word in ("vfd", "variable frequency", "variable speed", "inverter", "vsd"))
    bypass = "bypass" in method or "soft" in method
    sf_note = f", service factor {i.get('service_factor')}" if vfd else ""
    if vfd and bypass:
        code = "F" if high_sf else "B"
    elif vfd:
        code = "V" if high_sf else "D"
    elif "soft" in method:
        code = "S"
    else:
        return "", f"'{i['starting_method']}' is not a soft starter or VFD: choose"
    return code, f"Starting method '{i['starting_method']}'{sf_note}"


def _rule_hp(i, chosen, field):
    hp = _number(i.get("motor_hp"))
    if hp is None:
        return "", "No motor HP in the turn-in: choose"
    sizes = sorted((int(code), code) for code in field["options"])
    for size, code in sizes:
        if size >= hp:
            how = "matches" if size == hp else "next size up from"
            return code, f"{how} the {_text(hp)} HP motor"
    return "", f"{_text(hp)} HP is above the largest size ({sizes[-1][0]} HP): choose"


def _rule_line_voltage(i, chosen, field):
    volts, hz = _volts_hz(i.get("line_voltage"))
    if volts is None or hz is None:
        return "", "Voltage / Hz not readable: choose"
    best = None
    for code, description in field["options"].items():
        opt_volts, opt_hz = _volts_hz(description.replace("/", "V/"))
        if opt_hz == hz and opt_volts and abs(opt_volts - volts) / opt_volts <= 0.06:  # 480 -> 460, 600 -> 575
            if best is None or abs(opt_volts - volts) < best[0]:
                best = (abs(opt_volts - volts), code)
    if best is None:
        return "", f"No option for {_text(volts)} V / {_text(hz)} Hz: choose"
    return best[1], f"'{i['line_voltage']}'"


def _rule_control_voltage(i, chosen, field):
    volts, _ = _volts_hz(i.get("control_voltage"))
    if volts is None:
        return "", "No control voltage in the turn-in: choose"
    for code, description in field["options"].items():
        option = _number(description)
        if option and abs(option - volts) / option <= 0.10:   # 120 V -> 115 VAC, 240 V -> 230 VAC
            return code, f"Control panel '{i['control_voltage']}'"
    return "", f"No option for {_text(volts)} V control: choose"


def _rule_heater(i, chosen, field):
    answer = _yes_no(i.get("space_heaters"))
    if answer is None:
        return "", "Motor space heaters not stated: choose"
    return ("1" if answer else "0"), f"Motor space heaters: {i['space_heaters']}"


def _rule_motor_speed(i, chosen, field):
    rpm = _number(i.get("motor_speed"))
    if rpm is None:
        return "", "No motor speed in the turn-in: choose"
    _, hz = _volts_hz(i.get("line_voltage"))
    best = None
    for code, description in field["options"].items():
        sync = _number(description)
        if not sync:
            continue
        poles = 120 * hz / sync if hz else 2
        if poles != int(poles) or int(poles) % 2:   # speed not possible at this supply frequency
            continue
        if best is None or abs(sync - rpm) < best[0]:
            best = (abs(sync - rpm), code, sync)
    if best is None:
        return "", f"No synchronous speed fits {_text(rpm)} RPM: choose"
    hz_note = f" at {_text(hz)} Hz" if hz else ""
    return best[1], f"{_text(rpm)} RPM motor{hz_note} -> {_text(best[2])} RPM synchronous"


def _rule_reactor_filter(i, chosen, field):
    drive = chosen.get("Drive Type")
    if drive == "S":
        return "X", "Soft starter: not used"
    if drive:
        return "", "VFD: not in the turn-in, choose"
    return "", "Choose after Drive Type"


def _rule_ul_rating(i, chosen, field):
    text = _squash(i.get("starter_enclosure")).replace("type", "")
    if not text:
        return "", "No starter enclosure in the turn-in: choose"
    if not text.startswith("nema"):
        text = "nema" + text
    for code, description in field["options"].items():
        if _squash(description.rstrip("*")) == text:
            return code, f"Starter enclosure '{i['starter_enclosure']}'"
    return "", f"Enclosure '{i['starter_enclosure']}' not in the list: choose"


def _rule_not_in_turnin(what):
    return lambda i, chosen, field: ("", f"Not in the turn-in: choose {what}")


RULES = {
    "Compressor": _rule_compressor,
    "Drive Type": _rule_drive_type,
    "Frame": _rule_not_in_turnin("the disconnect type"),
    "HP": _rule_hp,
    "Line Voltage": _rule_line_voltage,
    "Control Voltage": _rule_control_voltage,
    "CPT": _rule_not_in_turnin("the control transformer size"),
    "Oil Pump": _rule_not_in_turnin("the oil pump motor (None if there is no pump)"),
    "Heater": _rule_heater,
    "Motor Speed": _rule_motor_speed,
    "Input Reactor": _rule_reactor_filter,
    "Output Filter": _rule_reactor_filter,
    "UL Rating": _rule_ul_rating,
}


def select_codes(inputs, table: dict, overrides: dict = None, pid_codes: dict = None) -> dict:
    """{field name: (code or "", reason)} for every field of the part-number table, in table order.
    `inputs` is {key: text} (see resolve_inputs); `overrides` is {field name: code} set by hand (kept as
    they are, and later rules see them). `pid_codes` is {field name: (code, where)} found on the P&ID: used when
    the turn-in does not decide a field, and noted in the reason when it disagrees with the turn-in.
    `inputs` may be a nested list (same-shape results back)."""
    if isinstance(inputs, (list, tuple)):
        return map_nested(lambda one: select_codes(one, table, overrides, pid_codes), inputs)
    values = {k: (v["value"] if isinstance(v, dict) else v) for k, v in inputs.items()}
    overrides = overrides or {}
    pid_codes = pid_codes or {}
    chosen, picks = {}, {}
    for field in table["fields"]:
        name = field["name"]
        if name in overrides:
            code = overrides[name]
            reason = "Set by hand"
        elif len(field["options"]) == 1:
            code, reason = next(iter(field["options"])), "Only option"
        else:
            rule = RULES.get(name, lambda i, c, f: ("", "No rule for this field: choose"))
            code, reason = rule(values, chosen, field)
            pid_code, where = pid_codes.get(name, ("", ""))
            if pid_code and not code:
                code, reason = pid_code, f"From the P&ID ({where})"
            elif pid_code and pid_code != code:
                reason += f"  [P&ID says {pid_code}, {where}]"
        chosen[name] = code
        picks[name] = (code, reason)
    return picks


def check_conditions(inputs) -> list:
    """Warnings for conditions that may need a derated or different starter: altitude above 3300 ft (1000 m),
    ambient above 104 °F (40 °C). `inputs` may be a nested list (same-shape results back)."""
    if isinstance(inputs, (list, tuple)):
        return map_nested(check_conditions, inputs)
    values = {k: (v["value"] if isinstance(v, dict) else v) for k, v in inputs.items()}
    warnings = []
    altitude = _number(values.get("altitude"))
    metres = "m" == _squash(values.get("altitude_units"))[:1] and "f" not in _squash(values.get("altitude_units"))
    limit = 1000 if metres else 3300
    if altitude is not None and altitude > limit:
        unit = values.get("altitude_units") or ("m" if metres else "ft")
        warnings.append(f"Altitude {_text(altitude)} {unit} is above {limit} {'m' if metres else 'ft'}: "
                        "check the starter's altitude derating.")
    temps = _numbers(values.get("ambient"))
    celsius = "c" in _squash(values.get("ambient_units")).replace("f", "")[:1]
    hot = 40 if celsius else 104
    if temps and max(temps) > hot:
        warnings.append(f"Ambient {values.get('ambient')} {values.get('ambient_units')} exceeds {hot}"
                        f"{'°C' if celsius else '°F'}: check the starter's temperature derating.")
    return warnings


# -------------------------------------------------------------------------------------------------- window
def _scrollable(parent):
    """A frame inside a vertically scrolling canvas. Returns (outer frame to place, inner frame to fill)."""
    outer = ttk.Frame(parent)
    background = ttk.Style().lookup("TFrame", "background") or None
    canvas = tk.Canvas(outer, highlightthickness=0, background=background)
    bar = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
    inner = ttk.Frame(canvas)
    inner.bind("<Configure>", lambda _e: canvas.configure(scrollregion=canvas.bbox("all")))
    window = canvas.create_window((0, 0), window=inner, anchor="nw")
    canvas.bind("<Configure>", lambda e: canvas.itemconfigure(window, width=e.width))
    canvas.configure(yscrollcommand=bar.set)
    canvas.pack(side="left", fill="both", expand=True)
    bar.pack(side="left", fill="y")
    wheel = lambda e: canvas.yview_scroll(int(-e.delta / 120), "units")
    outer.bind("<Enter>", lambda _e: canvas.bind_all("<MouseWheel>", wheel))
    outer.bind("<Leave>", lambda _e: canvas.unbind_all("<MouseWheel>"))
    return outer, inner


def run_starter_selection(values, pid_result=None, code_table=CODE_TABLE_FILE, save_path=None):
    """Window to check and edit the starter selection. Left: EVERY item extracted from the turn-in, grouped by
    sheet (editable; items used by the rules are marked and changing them re-runs the automatic selection).
    Right: every part-number field with its code (editable drop-down), the reason for it, and an "Auto" button to
    go back to the automatic choice. The part number at the top updates live.
    "Save" writes "<turn-in> - starter selection.json" in the Working folder (or `save_path`) and returns the
    selection dict; Cancel / closing returns None.
    `pid_result` is the submodule's P&ID instrument list ("<P&ID> - pids.json", dict or path, or None): every
    instrument it found is listed too, per page (value = Responsibility), for reference.
    `values` is the turn-in values JSON (dict or path); it may be a nested list (one window each, same shape back)."""
    if isinstance(values, (list, tuple)):
        return map_nested(lambda one: run_starter_selection(one, pid_result, code_table, save_path), values)
    values = _load_values(values)
    table = load_code_table(code_table)
    from_pid = pid_items(pid_result)
    items = extracted_items(values) + from_pid
    start = {item["key"]: item["value"] for item in items}
    groups = [(block["sheet"], block.get("status") or "") for block in values.get("sheets", [])]
    groups += [(sheet, "") for sheet in dict.fromkeys(i["sheet"] for i in from_pid)]
    instruments = sum(1 for i in from_pid if i.get("instrument"))
    if not pid_result:
        pid_note = ["No P&ID instrument list (the P&ID step was skipped or failed)."]
    elif not instruments:
        pid_note = ["The P&ID transcription found no instruments."]
    else:
        pid_note = []

    gc.collect()  # free Tk objects from earlier windows on this (main) thread, see GUI.run_with_progress
    root = make_window("Starter Selection", "1500x900", (1000, 600))
    style = ttk.Style()
    style.configure("Part.TLabel", font=("Consolas", 22, "bold"))
    style.configure("Head.TLabel", font=("Segoe UI", 10, "bold"))
    style.configure("Hint.TLabel", foreground="#9a9a9a")
    style.configure("Warn.TLabel", foreground="#ffb347")
    style.configure("Manual.TLabel", foreground="#6cb6ff")
    style.configure("Used.TLabel", foreground="#5fd38d")
    style.configure("Sheet.TLabel", font=("Segoe UI", 11, "bold"))

    outcome = {"result": None}
    input_vars = {key: tk.StringVar(value=text) for key, text in start.items()}
    option_text = {f["name"]: {code: f"{code} - {desc}" for code, desc in f["options"].items()}
                   for f in table["fields"]}
    field_vars = {f["name"]: tk.StringVar() for f in table["fields"]}
    reason_vars = {f["name"]: tk.StringVar() for f in table["fields"]}
    manual: set = set()           # fields set by hand: kept when inputs change
    part_var, warn_var = tk.StringVar(), tk.StringVar()
    reason_labels = {}

    def code_of(name: str) -> str:
        text = field_vars[name].get()
        return "" if text in ("", NOT_SELECTED) else text.split(" - ", 1)[0]

    def current_inputs() -> dict:
        return resolve_inputs({k: v.get() for k, v in input_vars.items()})

    def current_picks() -> dict:
        return select_codes(current_inputs(), table, {name: code_of(name) for name in manual})

    def refresh(*_args) -> None:
        picks = current_picks()
        for name, (code, reason) in picks.items():
            if name not in manual:
                field_vars[name].set(option_text[name].get(code, NOT_SELECTED))
            reason_vars[name].set(reason)
            reason_labels[name].configure(style="Manual.TLabel" if name in manual else
                                          ("Hint.TLabel" if code else "Warn.TLabel"))
        part_var.set(build_code_string(table, {n: c for n, (c, _) in picks.items()}))
        warnings = check_conditions(current_inputs()) + pid_note
        warn_var.set("\n".join(f"⚠ {w}" for w in warnings))

    # ---- top: part number and warnings ----
    top = ttk.Frame(root, padding=(16, 12, 16, 4))
    top.pack(fill="x")
    ttk.Label(top, text="Starter part number", style="Head.TLabel").pack(anchor="w")
    ttk.Label(top, textvariable=part_var, style="Part.TLabel").pack(anchor="w")
    ttk.Label(top, text="? = not selected yet. Orange reasons need a choice; blue = set by hand.",
              style="Hint.TLabel").pack(anchor="w")
    ttk.Label(top, textvariable=warn_var, style="Warn.TLabel", wraplength=1300, justify="left").pack(anchor="w")

    body = ttk.Frame(root, padding=(16, 4))
    body.pack(fill="both", expand=True)
    body.columnconfigure(0, weight=1)
    body.columnconfigure(1, weight=1)
    body.rowconfigure(0, weight=1)

    # ---- left: every extracted turn-in item, grouped by sheet ----
    left = ttk.LabelFrame(body, text="Extracted turn-in items (green = used for the selection; edit any value)",
                          padding=6)
    left.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
    outer, inputs_frame = _scrollable(left)
    outer.pack(fill="both", expand=True)
    inputs_frame.columnconfigure(1, weight=1)
    for col, head in enumerate(("Item", "Value", "Cell", "Used for")):
        ttk.Label(inputs_frame, text=head, style="Head.TLabel").grid(row=0, column=col, sticky="w", padx=4)
    row = 1
    for sheet, status in groups:
        ttk.Label(inputs_frame, text=sheet, style="Sheet.TLabel", wraplength=260, justify="left").grid(
            row=row, column=0, sticky="w", padx=4, pady=(10, 2))
        if status not in ("read", ""):
            ttk.Label(inputs_frame, text=status, style="Warn.TLabel").grid(
                row=row, column=1, columnspan=3, sticky="w", padx=4, pady=(10, 2))
        row += 1
        for item in (i for i in items if i["sheet"] == sheet):
            ttk.Label(inputs_frame, text=item["name"], style="Used.TLabel" if item["used_for"] else "TLabel",
                      wraplength=260, justify="left").grid(row=row, column=0, sticky="w", padx=(16, 4), pady=1)
            ttk.Entry(inputs_frame, textvariable=input_vars[item["key"]]).grid(
                row=row, column=1, sticky="ew", padx=4, pady=1)
            ttk.Label(inputs_frame, text=item["cell"], style="Hint.TLabel").grid(row=row, column=2, sticky="w", padx=4)
            ttk.Label(inputs_frame, text=item["used_for"], style="Used.TLabel").grid(
                row=row, column=3, sticky="w", padx=4)
            row += 1

    # ---- right: part-number fields ----
    right = ttk.LabelFrame(body, text="Part number fields (pick to set by hand, Auto to undo)", padding=6)
    right.grid(row=0, column=1, sticky="nsew")
    outer, fields_frame = _scrollable(right)
    outer.pack(fill="both", expand=True)
    fields_frame.columnconfigure(3, weight=1)
    for col, head in enumerate(("Pos", "Field", "Selection", "Reason", "")):
        ttk.Label(fields_frame, text=head, style="Head.TLabel").grid(row=0, column=col, sticky="w", padx=4)

    def on_pick(name: str) -> None:
        manual.add(name)
        refresh()

    def on_auto(name: str) -> None:
        manual.discard(name)
        refresh()

    for row, field in enumerate(table["fields"], start=1):
        name = field["name"]
        a, b = field["positions"]
        ttk.Label(fields_frame, text=f"{a}" if a == b else f"{a}-{b}", style="Hint.TLabel").grid(
            row=row, column=0, sticky="w", padx=4, pady=2)
        ttk.Label(fields_frame, text=name).grid(row=row, column=1, sticky="w", padx=4)
        box = ttk.Combobox(fields_frame, textvariable=field_vars[name], state="readonly", width=40,
                           values=[NOT_SELECTED] + list(option_text[name].values()))
        box.grid(row=row, column=2, sticky="w", padx=4, pady=2)
        box.bind("<<ComboboxSelected>>", lambda _e, n=name: on_pick(n))
        reason_labels[name] = ttk.Label(fields_frame, textvariable=reason_vars[name], wraplength=380,
                                        justify="left")
        reason_labels[name].grid(row=row, column=3, sticky="w", padx=4)
        ttk.Button(fields_frame, text="Auto", width=6, command=lambda n=name: on_auto(n)).grid(
            row=row, column=4, padx=4)

    # ---- bottom: buttons ----
    def reset_all() -> None:
        manual.clear()
        for key, text in start.items():
            input_vars[key].set(text)
        refresh()

    def save() -> None:
        picks = current_picks()
        part = build_code_string(table, {n: c for n, (c, _) in picks.items()})
        if "?" in part and not messagebox.askyesno(
                "Incomplete", f"{part}\n\nSome fields are not selected (?). Save anyway?"):
            return
        source = values.get("source_workbook")
        stem = Path(source).stem if source else "turn-in"
        target = Path(save_path) if save_path else working_path(f"{stem}{SELECTION_FILE_SUFFIX}.json")
        result = {
            "part_number": part,
            "complete": "?" not in part,
            "source_workbook": source,
            "saved": datetime.datetime.now().isoformat(timespec="seconds"),
            "warnings": check_conditions(current_inputs()) + pid_note,
            "pid_result": str(pid_result) if pid_result and not isinstance(pid_result, dict) else None,
            "inputs_used": {spec["label"]: current_inputs()[spec["key"]] for spec in INPUTS},
            "items": [{"sheet": i["sheet"], "name": i["name"], "cell": i["cell"],
                       "value": input_vars[i["key"]].get(), "turn_in_value": i["value"],
                       "edited": input_vars[i["key"]].get() != i["value"], "used_for": i["used_for"]}
                      for i in items],
            "fields": [{"name": f["name"], "positions": f["positions"], "code": picks[f["name"]][0],
                        "description": f["options"].get(picks[f["name"]][0], ""),
                        "set_by": "hand" if f["name"] in manual else "automatic",
                        "reason": picks[f["name"]][1]} for f in table["fields"]],
        }
        try:
            target.write_text(json.dumps(result, indent=2, ensure_ascii=False), encoding="utf-8")
        except OSError as exc:
            messagebox.showerror("Error", f"Could not save:\n{target}\n\n{exc}")
            return
        result["saved_to"] = str(target)
        outcome["result"] = result
        root.destroy()

    bar = ttk.Frame(root, padding=(16, 8, 16, 14))
    bar.pack(side="bottom", fill="x", before=body)  # pinned to the bottom so it stays visible at any window size
    ttk.Button(bar, text="Reset all to turn-in / automatic", command=reset_all).pack(side="left")
    ttk.Button(bar, text="Save", width=10, command=save).pack(side="right")
    ttk.Button(bar, text="Cancel", width=10, command=root.destroy).pack(side="right", padx=(0, 8))

    for var in input_vars.values():
        var.trace_add("write", refresh)
    refresh()
    root.protocol("WM_DELETE_WINDOW", root.destroy)
    root.mainloop()

    # Window is gone: drop the Tk references here, on the main thread (see GUI.run_with_progress).
    del root, style, input_vars, field_vars, reason_vars, part_var, warn_var, reason_labels
    gc.collect()
    return outcome["result"]
