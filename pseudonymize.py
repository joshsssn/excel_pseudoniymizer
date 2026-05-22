"""
pseudo_excel.py — Pseudonymiseur Excel local, multi-feuilles, GUI conviviale
    pip install pandas openpyxl python-dateutil
"""
import copy
import json
import re
import secrets
import hashlib
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd


# ============================================================
#  MOTEUR
# ============================================================

STRATEGIES = {
    "keep":               "Ne rien faire (garder tel quel)",
    "sequential_id":      "Remplacer par ID_1, ID_2… (mêmes valeurs → même ID)",
    "scale_numeric":      "Multiplier par un facteur (préserve les ratios)",
    "regex_replace":      "Remplacer une partie via motif (ex: chiffres d'un libellé)",
    "hash_deterministic": "Empreinte stable (hash) — utile pour jointures",
    "date_shift":         "Décaler toutes les dates d'un nombre de jours",
    "recalc_formula":     "Recalculer après pseudo (ex: TTC = HT*(1+TVA-rate))",
}

# Paramètres par stratégie : (nom_param, label_humain, default, type)
STRATEGY_PARAMS = {
    "keep": [],
    "sequential_id": [
        ("prefix", "Préfixe à utiliser", "ID", str),
    ],
    "scale_numeric": [
        ("factor", "Facteur multiplicateur (laisser vide = aléatoire 0.5–2.0)", "", float),
    ],
    "regex_replace": [
        ("pattern", "Motif (regex) à remplacer", r"\d+", str),
        ("prefix", "Préfixe pour les remplacements", "N", str),
    ],
    "hash_deterministic": [
        ("length", "Longueur de l'empreinte (caractères)", 10, int),
    ],
    "date_shift": [
        ("days", "Décalage en jours (vide = aléatoire ±1000)", "", int),
    ],
    "recalc_formula": [
        ("formula", "Formule (utilisez les noms de colonnes, ex: HT*(1+`TVA-rate`))", "", str),
    ],
}

# ============================================================
#  HELPERS NaN / TYPES — à ajouter en haut du fichier
# ============================================================

# Sentinelle : valeurs qu'on considère comme "vides"
def is_missing(v) -> bool:
    """Détecte tous les types de valeurs vides : None, NaN, NaT, pd.NA, '', 'nan', 'NaT'."""
    if v is None:
        return True
    try:
        if pd.isna(v):
            return True
    except (TypeError, ValueError):
        pass
    if isinstance(v, str) and v.strip().lower() in ("", "nan", "nat", "none", "<na>", "null"):
        return True
    return False


def coerce_numeric_smart(series: pd.Series) -> tuple[pd.Series, int]:
    """
    Convertit en numérique en gérant les formats FR/EN courants.
    Retourne (série_numérique, nb_valeurs_non_parseables).
    """
    if pd.api.types.is_numeric_dtype(series):
        return series.astype(float), 0

    def clean(v):
        if is_missing(v):
            return None
        s = str(v).strip()
        # retire symboles monétaires et espaces (y compris insécables)
        s = re.sub(r"[€$£¥\s\u00a0]", "", s)
        if not s:
            return None
        # heuristique FR vs EN : si une virgule ET un point → le dernier des deux est le séparateur décimal
        if "," in s and "." in s:
            if s.rfind(",") > s.rfind("."):
                s = s.replace(".", "").replace(",", ".")
            else:
                s = s.replace(",", "")
        elif "," in s:
            # virgule seule → décimale FR
            s = s.replace(",", ".")
        try:
            return float(s)
        except ValueError:
            return None

    cleaned = series.map(clean)
    n_failed = cleaned.isna().sum() - series.map(is_missing).sum()
    return pd.to_numeric(cleaned, errors="coerce"), int(max(0, n_failed))


class PseudoEngine:
    """
    Map structure :
      {
        "_shared_lookups": { "id_client": {src: pseudo, ...} },
        "_date_formats":   { "sheet::col": "%d/%m/%Y" },
        "sheets": {
            "Feuille1": {
                "col_a": {"strategy": "...", "params": {...}, "lookup_key": "id_client"}
            }
        }
      }
    """

    def __init__(self, existing: dict | None = None):
        if existing:
            self.map = existing
            # garantit la structure même sur vieilles maps
            self.map.setdefault("_shared_lookups", {})
            self.map.setdefault("_date_formats", {})
            self.map.setdefault("sheets", {})
        else:
            self.map = {"_shared_lookups": {}, "_date_formats": {}, "sheets": {}}
        self._warnings: list[str] = []

    # ---- helpers ----

    def _get_lookup(self, col_name: str) -> dict:
        """Lookup partagé par nom de colonne (jointures entre feuilles)."""
        return self.map["_shared_lookups"].setdefault(col_name, {})

    # ---- stratégies ----

    def _seq_id(self, series, col, prefix):
        lookup = self._get_lookup(col)
        out = []
        for v in series:
            if is_missing(v):
                out.append(None); continue
            key = str(v).strip()
            if key not in lookup:
                lookup[key] = f"{prefix}_{len(lookup) + 1}"
            out.append(lookup[key])
        return pd.Series(out, index=series.index, dtype="object")

    def _scale(self, series, col, factor):
        if factor in (None, "", "None"):
            factor = round(0.5 + secrets.randbelow(1500) / 1000, 4)
        factor = float(factor)
        self.map["_shared_lookups"].setdefault(f"__factor__{col}", {"value": factor})

        numeric, n_failed = coerce_numeric_smart(series)
        if n_failed > 0:
            self._warnings.append(
                f"Colonne '{col}' : {n_failed} valeur(s) non numérique(s) → conservées telles quelles."
            )
            # on garde l'original là où le parsing a échoué
            result = series.copy().astype(object)
            mask_ok = numeric.notna()
            result[mask_ok] = (numeric[mask_ok] * factor)
            # on remet les NaN d'origine
            result[series.map(is_missing)] = None
            return result
        return numeric * factor

    def _regex(self, series, col, pattern, prefix):
        lookup = self._get_lookup(col)
        try:
            rx = re.compile(pattern)
        except re.error as e:
            raise ValueError(f"Motif regex invalide pour '{col}' : {e}")

        def repl(m):
            captured = m.group(0)
            if captured not in lookup:
                lookup[captured] = f"{prefix}_{len(lookup) + 1}"
            return lookup[captured]

        def transform(v):
            if is_missing(v): return None
            return rx.sub(repl, str(v))
        return series.map(transform)

    def _hash(self, series, col, length):
        lookup = self._get_lookup(col)
        length = int(length)
        out = []
        for v in series:
            if is_missing(v):
                out.append(None); continue
            key = str(v).strip()
            if key not in lookup:
                lookup[key] = hashlib.sha256(key.encode("utf-8")).hexdigest()[:length]
            out.append(lookup[key])
        return pd.Series(out, index=series.index, dtype="object")

    def _date_shift(self, series, col, sheet, days):
        if days in (None, "", "None"):
            days = secrets.randbelow(2000) - 1000
        days = int(days)
        self.map["_shared_lookups"].setdefault(f"__date_shift__{col}", {"days": days})

        fmt = detect_date_format(series)
        self.map["_date_formats"][f"{sheet}::{col}"] = fmt

        parsed = parse_dates_robust(series, hint_format=fmt)
        n_failed = parsed.isna().sum() - series.map(is_missing).sum()
        if n_failed > 0:
            self._warnings.append(
                f"Colonne '{col}' : {n_failed} date(s) non reconnue(s) → conservées telles quelles."
            )

        shifted = parsed + pd.Timedelta(days=days)

        # Reconstruction propre : original si non-parseable, vide si missing, formaté sinon
        result = []
        for orig, sh in zip(series, shifted):
            if is_missing(orig):
                result.append(None)
            elif pd.isna(sh):
                result.append(orig)  # garde la valeur d'origine si parsing échoué
            else:
                result.append(sh.strftime(fmt) if fmt else sh)
        return pd.Series(result, index=series.index, dtype="object")

    # ---- pipeline ----

    def apply_sheet(self, df: pd.DataFrame, sheet: str, config: dict) -> pd.DataFrame:
        out = df.copy()
        sheet_map = self.map["sheets"].setdefault(sheet, {})

        # 1er passage
        for col, cfg in config.items():
            strat, p = cfg["strategy"], cfg.get("params", {})
            sheet_map[col] = {"strategy": strat, "params": p, "lookup_key": col}

            if strat == "keep" or strat == "recalc_formula":
                continue
            elif strat == "sequential_id":
                out[col] = self._seq_id(out[col], col, p.get("prefix", "ID"))
            elif strat == "scale_numeric":
                out[col] = self._scale(out[col], col, p.get("factor"))
            elif strat == "regex_replace":
                out[col] = self._regex(out[col], col, p.get("pattern", r"\d+"), p.get("prefix", "N"))
            elif strat == "hash_deterministic":
                out[col] = self._hash(out[col], col, p.get("length", 10))
            elif strat == "date_shift":
                out[col] = self._date_shift(out[col], col, sheet, p.get("days"))

        # 2e passage : recalculs
        for col, cfg in config.items():
            if cfg["strategy"] == "recalc_formula":
                formula = cfg["params"].get("formula", "")
                if formula.strip():
                    out[col] = eval_formula(formula, out)
        return out

    def apply_workbook(self, sheets: dict[str, pd.DataFrame], configs: dict) -> dict:
        return {name: self.apply_sheet(df, name, configs.get(name, {}))
                for name, df in sheets.items()}

    # ---- inverse ----

    def reverse_sheet(self, df: pd.DataFrame, sheet: str) -> pd.DataFrame:
        out = df.copy()
        sheet_map = self.map["sheets"].get(sheet, {})
        for col, entry in sheet_map.items():
            if col not in out.columns: continue
            strat = entry["strategy"]
            if strat in ("keep", "recalc_formula"):
                continue

            if strat in ("sequential_id", "hash_deterministic"):
                lookup = self._get_lookup(col)
                inv = {v: k for k, v in lookup.items()}
                out[col] = out[col].map(
                    lambda v: inv.get(v, v) if not is_missing(v) else None
                )

            elif strat == "regex_replace":
                lookup = self._get_lookup(col)
                inv = {v: k for k, v in lookup.items()}
                if inv:
                    rx = re.compile("|".join(re.escape(v) for v in inv))
                    def revert(v):
                        if is_missing(v): return None
                        return rx.sub(lambda m: inv[m.group(0)], str(v))
                    out[col] = out[col].map(revert)

            elif strat == "scale_numeric":
                key = f"__factor__{col}"
                shared = self.map["_shared_lookups"].get(key)
                if not shared or "value" not in shared:
                    self._warnings.append(
                        f"Colonne '{col}' : facteur d'échelle introuvable dans la map → ignorée."
                    )
                    continue
                factor = float(shared["value"])
                numeric, _ = coerce_numeric_smart(out[col])
                result = out[col].copy().astype(object)
                mask = numeric.notna()
                result[mask] = numeric[mask] / factor
                result[out[col].map(is_missing)] = None
                out[col] = result

            elif strat == "date_shift":
                key = f"__date_shift__{col}"
                shared = self.map["_shared_lookups"].get(key)
                if not shared or "days" not in shared:
                    self._warnings.append(
                        f"Colonne '{col}' : décalage de date introuvable → ignorée."
                    )
                    continue
                days = int(shared["days"])
                fmt = self.map["_date_formats"].get(f"{sheet}::{col}")
                parsed = parse_dates_robust(out[col], hint_format=fmt)
                back = parsed - pd.Timedelta(days=days)
                result = []
                for orig, b in zip(out[col], back):
                    if is_missing(orig):
                        result.append(None)
                    elif pd.isna(b):
                        result.append(orig)
                    else:
                        result.append(b.strftime(fmt) if fmt else b)
                out[col] = pd.Series(result, index=out[col].index, dtype="object")

        return out

    def reverse_workbook(self, sheets: dict[str, pd.DataFrame]) -> dict:
        return {name: self.reverse_sheet(df, name) for name, df in sheets.items()}

    def save(self, path: Path):
        path.write_text(json.dumps(self.map, indent=2, ensure_ascii=False, default=str),
                        encoding="utf-8")

    @classmethod
    def load(cls, path: Path) -> "PseudoEngine":
        return cls(json.loads(path.read_text(encoding="utf-8")))


# ============================================================
#  UTILITAIRES (dates + formules)
# ============================================================

DATE_FORMATS_TRY = [
    "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y",
    "%Y-%m-%d", "%Y/%m/%d",
    "%m/%d/%Y", "%d/%m/%Y %H:%M", "%Y-%m-%d %H:%M:%S",
    "%d %B %Y", "%d %b %Y",
]

def detect_date_format(series: pd.Series) -> str | None:
    samples = [str(v) for v in series.dropna().head(20) if not is_missing(v)]
    if not samples: return None
    best, best_score = None, 0
    for fmt in DATE_FORMATS_TRY:
        score = 0
        for s in samples:
            try:
                pd.to_datetime(s, format=fmt); score += 1
            except (ValueError, TypeError):
                pass
        if score > best_score:
            best, best_score = fmt, score
    return best if best_score >= len(samples) * 0.5 else None


def parse_dates_robust(series: pd.Series, hint_format: str | None = None) -> pd.Series:
    """Parse les dates ; si un format est connu, on l'utilise en priorité."""
    if hint_format:
        try:
            return pd.to_datetime(series, format=hint_format, errors="coerce")
        except Exception:
            pass
    # heuristique dayfirst : True si format détecté commence par %d
    fmt = detect_date_format(series)
    dayfirst = bool(fmt and fmt.startswith("%d"))
    try:
        return pd.to_datetime(series, errors="coerce", dayfirst=dayfirst)
    except Exception:
        return pd.Series([pd.NaT] * len(series), index=series.index)

def eval_formula(formula: str, df: pd.DataFrame) -> pd.Series:
    """Évalue une formule avec substitution sûre par placeholders uniques."""
    placeholders: dict[str, str] = {}
    used_cols: list[str] = []
    expr = formula

    # 1. Backticks d'abord (priorité absolue, names littéraux)
    for c in df.columns:
        token = f"`{c}`"
        if token in expr:
            ph = f"__COL_{len(placeholders)}__"
            placeholders[ph] = str(c)
            expr = expr.replace(token, ph)
            used_cols.append(c)

    # 2. Noms simples, ordre décroissant pour éviter prefix-match
    for c in sorted([str(x) for x in df.columns], key=len, reverse=True):
        if any(p == c for p in placeholders.values()): continue
        # mot entier, pas dans un placeholder existant
        pattern = rf"(?<![A-Za-z0-9_]){re.escape(c)}(?![A-Za-z0-9_])"
        new_expr, n = re.subn(pattern, lambda m, col=c: _make_ph(placeholders, col), expr)
        if n > 0:
            expr = new_expr
            used_cols.append(c)

    # 3. Remplacement final placeholders → __df__['col']
    for ph, col in placeholders.items():
        expr = expr.replace(ph, f"__df__[{col!r}]")

    # Conversion numérique défensive sur colonnes utilisées
    safe_df = df.copy()
    for c in set(used_cols):
        if c in safe_df.columns and not pd.api.types.is_numeric_dtype(safe_df[c]):
            safe_df[c], _ = coerce_numeric_smart(safe_df[c])

    try:
        return eval(expr, {"__builtins__": {}}, {"__df__": safe_df})
    except KeyError as e:
        raise ValueError(f"Formule fait référence à une colonne inexistante : {e}")
    except Exception as e:
        raise ValueError(f"Erreur dans la formule '{formula}' : {e}")


def _make_ph(placeholders: dict, col: str) -> str:
    # réutilise un placeholder existant pour la même colonne
    for ph, c in placeholders.items():
        if c == col: return ph
    ph = f"__COL_{len(placeholders)}__"
    placeholders[ph] = col
    return ph


# ============================================================
#  STYLE (thème JetBrains-like, sombre doux)
# ============================================================

THEME = {
    "bg":          "#2b2d30",   # fond principal
    "bg_alt":      "#1e1f22",   # fond secondaire (cartes, inputs)
    "bg_hover":    "#3a3d41",
    "border":      "#3e4045",
    "fg":          "#dfe1e5",   # texte principal
    "fg_dim":      "#8b8d91",   # texte secondaire
    "fg_mute":     "#6c6e72",
    "accent":      "#5394ec",   # bleu actions
    "accent_hov":  "#6aa3f0",
    "success":     "#5fb865",
    "warning":     "#e8a33d",
    "danger":      "#e8635a",
    "selection":   "#2e436e",
}

FONT_BASE   = ("Segoe UI", 10)
FONT_SMALL  = ("Segoe UI", 9)
FONT_BOLD   = ("Segoe UI", 10, "bold")
FONT_TITLE  = ("Segoe UI", 12, "bold")
FONT_MONO   = ("Consolas", 9)


def apply_theme(root: tk.Tk):
    style = ttk.Style(root)
    style.theme_use("clam")
    t = THEME

    root.configure(bg=t["bg"])

    # Frames
    style.configure("TFrame", background=t["bg"])
    style.configure("Card.TFrame", background=t["bg_alt"], relief="flat")
    style.configure("Toolbar.TFrame", background=t["bg_alt"])

    # Labels
    style.configure("TLabel", background=t["bg"], foreground=t["fg"], font=FONT_BASE)
    style.configure("Card.TLabel", background=t["bg_alt"], foreground=t["fg"], font=FONT_BASE)
    style.configure("Title.TLabel", background=t["bg"], foreground=t["fg"], font=FONT_TITLE)
    style.configure("Dim.TLabel", background=t["bg_alt"], foreground=t["fg_dim"], font=FONT_SMALL)
    style.configure("Mute.TLabel", background=t["bg"], foreground=t["fg_mute"], font=FONT_SMALL)
    style.configure("Status.TLabel", background=t["bg_alt"], foreground=t["fg_dim"],
                    font=FONT_SMALL, padding=6)
    style.configure("Warn.TLabel", background="#3a2f1f", foreground=t["warning"],
                    font=FONT_SMALL, padding=8)

    # Boutons
    style.configure("TButton",
        background=t["bg_alt"], foreground=t["fg"],
        bordercolor=t["border"], lightcolor=t["bg_alt"], darkcolor=t["bg_alt"],
        focusthickness=0, padding=(12, 6), font=FONT_BASE, relief="flat")
    style.map("TButton",
        background=[("active", t["bg_hover"]), ("pressed", t["bg_hover"])],
        foreground=[("active", t["fg"])])

    style.configure("Accent.TButton",
        background=t["accent"], foreground="#ffffff",
        bordercolor=t["accent"], padding=(14, 7), font=FONT_BOLD, relief="flat")
    style.map("Accent.TButton",
        background=[("active", t["accent_hov"]), ("pressed", t["accent_hov"])])

    style.configure("Ghost.TButton",
        background=t["bg"], foreground=t["fg_dim"],
        padding=(8, 4), font=FONT_SMALL, relief="flat")
    style.map("Ghost.TButton",
        background=[("active", t["bg_alt"])],
        foreground=[("active", t["fg"])])

    # Combobox
    style.configure("TCombobox",
        fieldbackground=t["bg_alt"], background=t["bg_alt"],
        foreground=t["fg"], bordercolor=t["border"],
        arrowcolor=t["fg_dim"], padding=4)
    style.map("TCombobox",
        fieldbackground=[("readonly", t["bg_alt"])],
        foreground=[("readonly", t["fg"])])
    root.option_add("*TCombobox*Listbox.background", t["bg_alt"])
    root.option_add("*TCombobox*Listbox.foreground", t["fg"])
    root.option_add("*TCombobox*Listbox.selectBackground", t["selection"])
    root.option_add("*TCombobox*Listbox.font", FONT_BASE)

    # Entry
    style.configure("TEntry",
        fieldbackground=t["bg_alt"], foreground=t["fg"],
        bordercolor=t["border"], insertcolor=t["fg"], padding=4)

    # Notebook
    style.configure("TNotebook", background=t["bg"], borderwidth=0, tabmargins=(8, 6, 0, 0))
    style.configure("TNotebook.Tab",
        background=t["bg"], foreground=t["fg_dim"],
        padding=(16, 8), font=FONT_BASE, borderwidth=0)
    style.map("TNotebook.Tab",
        background=[("selected", t["bg_alt"])],
        foreground=[("selected", t["fg"])])

    # Treeview
    style.configure("Treeview",
        background=t["bg_alt"], fieldbackground=t["bg_alt"],
        foreground=t["fg"], bordercolor=t["border"],
        rowheight=24, font=FONT_BASE)
    style.configure("Treeview.Heading",
        background=t["bg"], foreground=t["fg_dim"],
        font=FONT_BOLD, relief="flat", padding=6)
    style.map("Treeview",
        background=[("selected", t["selection"])],
        foreground=[("selected", t["fg"])])
    style.map("Treeview.Heading",
        background=[("active", t["bg_hover"])])

    # Scrollbar
    style.configure("Vertical.TScrollbar",
        background=t["bg"], troughcolor=t["bg"],
        bordercolor=t["bg"], arrowcolor=t["fg_dim"], gripcount=0)
    style.configure("Horizontal.TScrollbar",
        background=t["bg"], troughcolor=t["bg"],
        bordercolor=t["bg"], arrowcolor=t["fg_dim"], gripcount=0)


# ============================================================
#  WIDGETS UTILITAIRES
# ============================================================

class Tooltip:
    """Tooltip simple au survol."""
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tip = None
        widget.bind("<Enter>", self._show)
        widget.bind("<Leave>", self._hide)

    def _show(self, _):
        if self.tip or not self.text: return
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 4
        self.tip = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        lbl = tk.Label(tw, text=self.text, justify="left",
                       background=THEME["bg_alt"], foreground=THEME["fg"],
                       relief="solid", borderwidth=1, font=FONT_SMALL,
                       padx=8, pady=4, wraplength=320)
        lbl.pack()

    def _hide(self, _):
        if self.tip:
            self.tip.destroy()
            self.tip = None


class ColumnCard(ttk.Frame):
    """Carte de configuration pour une colonne."""

    STATUS_ICONS = {
        "keep":               ("○", THEME["fg_mute"]),
        "sequential_id":      ("●", THEME["accent"]),
        "scale_numeric":      ("●", THEME["accent"]),
        "regex_replace":      ("●", THEME["accent"]),
        "hash_deterministic": ("●", THEME["accent"]),
        "date_shift":         ("●", THEME["accent"]),
        "recalc_formula":     ("∑", THEME["success"]),
    }

    def __init__(self, parent, tab, col_name, df_col, on_change):
        super().__init__(parent, style="Card.TFrame", padding=12)
        self.tab = tab
        self.col_name = col_name
        self.on_change = on_change
        self.param_vars: dict = {}

        # Border émulée par un canvas-frame
        self.configure(borderwidth=1, relief="solid")

        # --- Ligne 1 : statut + nom + stratégie ---
        head = ttk.Frame(self, style="Card.TFrame")
        head.pack(fill="x")

        self.status_lbl = tk.Label(head, text="○", bg=THEME["bg_alt"],
                                   fg=THEME["fg_mute"], font=("Segoe UI", 14))
        self.status_lbl.pack(side="left", padx=(0, 10))

        name_box = ttk.Frame(head, style="Card.TFrame")
        name_box.pack(side="left", fill="x", expand=True)

        ttk.Label(name_box, text=str(col_name),
                  style="Card.TLabel", font=FONT_BOLD).pack(anchor="w")

        preview = ", ".join(str(v) for v in df_col.head(3).tolist() if not is_missing(v))[:60]
        if not preview: preview = "(toutes les valeurs sont vides)"
        ttk.Label(name_box, text=preview, style="Dim.TLabel").pack(anchor="w")

        right = ttk.Frame(head, style="Card.TFrame")
        right.pack(side="right")

        self.strat_var = tk.StringVar()
        self.strat = ttk.Combobox(right, textvariable=self.strat_var,
            values=list(STRATEGIES.keys()), state="readonly", width=22)
        self.strat.pack(side="left", padx=(0, 6))

        info_btn = ttk.Button(right, text="ⓘ", style="Ghost.TButton", width=3,
            command=self._show_help)
        info_btn.pack(side="left")
        Tooltip(info_btn, "Voir l'explication de la stratégie")

        # --- Ligne 2 : paramètres (initialement vide) ---
        self.params_frame = ttk.Frame(self, style="Card.TFrame")
        self.params_frame.pack(fill="x", pady=(10, 0))

        self.strat_var.trace_add("write", lambda *_: self._on_strategy_change())

    def set_strategy(self, name: str):
        self.strat_var.set(name)

    def _on_strategy_change(self):
        for child in self.params_frame.winfo_children():
            child.destroy()
        self.param_vars.clear()

        strat = self.strat_var.get()
        icon, color = self.STATUS_ICONS.get(strat, ("○", THEME["fg_mute"]))
        self.status_lbl.config(text=icon, fg=color)

        params = STRATEGY_PARAMS.get(strat, [])
        if not params:
            ttk.Label(self.params_frame,
                text=STRATEGIES.get(strat, ""),
                style="Dim.TLabel").pack(anchor="w")
        else:
            defaults = self.tab._smart_defaults(self.col_name, strat)
            for pname, plabel, pdefault, ptype in params:
                row = ttk.Frame(self.params_frame, style="Card.TFrame")
                row.pack(fill="x", pady=2)
                ttk.Label(row, text=plabel, style="Dim.TLabel",
                          width=42, anchor="w").pack(side="left")
                var = tk.StringVar(value=str(defaults.get(pname, pdefault)))
                entry = ttk.Entry(row, textvariable=var)
                entry.pack(side="left", fill="x", expand=True)
                var.trace_add("write", lambda *_: self.on_change())
                self.param_vars[pname] = (var, ptype)

        self.on_change()

    def _show_help(self):
        s = self.strat_var.get()
        if s in STRATEGIES:
            messagebox.showinfo("Aide — " + s, STRATEGIES[s])

    def get_config(self) -> dict:
        params = {}
        for pname, (var, ptype) in self.param_vars.items():
            raw = var.get().strip()
            if raw == "":
                params[pname] = ""
            else:
                try: params[pname] = ptype(raw)
                except (ValueError, TypeError): params[pname] = raw
        return {"strategy": self.strat_var.get(), "params": params}


# ============================================================
#  ONGLET FEUILLE
# ============================================================

class SheetTab(ttk.Frame):

    def __init__(self, parent, app, sheet_name, df):
        super().__init__(parent)
        self.app = app
        self.sheet_name = sheet_name
        self.df = df
        self.cards: dict[str, ColumnCard] = {}

        # Layout : panneau gauche (config) | panneau droit (preview)
        paned = ttk.PanedWindow(self, orient="horizontal")
        paned.pack(fill="both", expand=True, padx=8, pady=8)

        # --- gauche : config colonnes ---
        left = ttk.Frame(paned)
        paned.add(left, weight=3)

        header_left = ttk.Frame(left)
        header_left.pack(fill="x", pady=(0, 8))
        ttk.Label(header_left,
            text=f"Configuration — {sheet_name}",
            style="Title.TLabel").pack(side="left")
        ttk.Label(header_left,
            text=f"  {len(df.columns)} colonnes · {len(df)} lignes",
            style="Mute.TLabel").pack(side="left")

        # zone scrollable de cartes
        self._build_scrollable_cards(left)

        # --- droite : aperçu ---
        right = ttk.Frame(paned)
        paned.add(right, weight=2)

        header_right = ttk.Frame(right)
        header_right.pack(fill="x", pady=(0, 8))
        ttk.Label(header_right, text="Aperçu", style="Title.TLabel").pack(side="left")
        ttk.Button(header_right, text="↻", style="Ghost.TButton",
                   command=self.refresh_preview).pack(side="right")

        self.warn_lbl = ttk.Label(right, text="", style="Warn.TLabel")
        # affiché seulement s'il y a des warnings

        prev_box = ttk.Frame(right)
        prev_box.pack(fill="both", expand=True)

        ttk.Label(prev_box, text="ORIGINAL", style="Mute.TLabel").pack(anchor="w", pady=(4, 2))
        self.tree_orig = self._make_tree(prev_box)
        self.tree_orig.pack(fill="both", expand=True)

        ttk.Label(prev_box, text="PSEUDONYMISÉ", style="Mute.TLabel").pack(anchor="w", pady=(10, 2))
        self.tree_new = self._make_tree(prev_box)
        self.tree_new.pack(fill="both", expand=True)

        self.refresh_preview()

    def _build_scrollable_cards(self, parent):
        wrap = ttk.Frame(parent)
        wrap.pack(fill="both", expand=True)

        canvas = tk.Canvas(wrap, bg=THEME["bg"], highlightthickness=0)
        sb = ttk.Scrollbar(wrap, orient="vertical", command=canvas.yview)
        cards_frame = ttk.Frame(canvas)

        cards_frame.bind("<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas_window = canvas.create_window((0, 0), window=cards_frame, anchor="nw")
        canvas.bind("<Configure>",
            lambda e: canvas.itemconfig(canvas_window, width=e.width))
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        # scroll molette
        canvas.bind_all("<MouseWheel>",
            lambda e: canvas.yview_scroll(int(-e.delta/120), "units"), add="+")

        for col in self.df.columns:
            card = ColumnCard(cards_frame, self, col, self.df[col],
                              on_change=self.app.schedule_preview)
            card.pack(fill="x", pady=4, padx=4)
            self.cards[col] = card
            card.set_strategy(self._suggest_strategy(col))

    def _make_tree(self, parent) -> ttk.Treeview:
        cols = [str(c) for c in self.df.columns]
        tree = ttk.Treeview(parent, columns=cols, show="headings", height=6)
        for c in cols:
            tree.heading(c, text=c)
            tree.column(c, width=110, anchor="w", stretch=True)
        return tree

    def _suggest_strategy(self, col: str) -> str:
        cl = str(col).lower()
        tokens = re.split(r"[\s_\-./]+", cl)
        has_token = lambda *kws: any(t in kws for t in tokens)
        if "ttc" in tokens: return "recalc_formula"
        if has_token("tva", "rate", "taux", "vat"): return "keep"
        if has_token("ht", "montant", "prix", "amount", "price"): return "scale_numeric"
        if has_token("id", "client", "nom", "name", "customer"): return "sequential_id"
        if has_token("date"): return "date_shift"
        if has_token("libelle", "libellé", "ref", "label", "reference"): return "regex_replace"
        return "keep"

    def _smart_defaults(self, col: str, strat: str) -> dict:
        cl = str(col).lower()
        tokens = re.split(r"[\s_\-./]+", cl)
        has_token = lambda *kws: any(t in kws for t in tokens)
        if strat == "sequential_id":
            if has_token("client"): return {"prefix": "CLI"}
            if has_token("nom", "name"): return {"prefix": "NOM"}
            if has_token("id"): return {"prefix": "ID"}
            return {"prefix": str(col).upper()[:5]}
        if strat == "recalc_formula" and "ttc" in tokens:
            cols = list(self.df.columns)
            ht = next(
                (c for c in cols
                 if "ht" in re.split(r"[\s_\-./]+", str(c).lower())),
                "HT")
            tva = next(
                (c for c in cols
                 if any(k in re.split(r"[\s_\-./]+", str(c).lower())
                        for k in ("tva", "rate", "taux", "vat"))),
                "TVA-rate")
            return {"formula": f"`{ht}`*(1+`{tva}`)"}
        if strat == "regex_replace":
            return {"pattern": r"\d+", "prefix": "N"}
        return {}

    def get_config(self) -> dict:
        return {col: card.get_config() for col, card in self.cards.items()}

    def refresh_preview(self):
        try:
            existing = copy.deepcopy(self.app.existing_map) if self.app.existing_map else None
            engine = PseudoEngine(existing)
            cfg = self.get_config()
            preview_df = engine.apply_sheet(self.df.head(8).copy(), self.sheet_name, cfg)
            warnings = engine._warnings
        except Exception as e:
            self._fill_tree(self.tree_orig, self.df.head(8))
            self.tree_new.delete(*self.tree_new.get_children())
            self.warn_lbl.config(text=f"⚠ Aperçu indisponible : {e}")
            self.warn_lbl.pack(fill="x", pady=4) if not self.warn_lbl.winfo_ismapped() else None
            return

        self._fill_tree(self.tree_orig, self.df.head(8))
        self._fill_tree(self.tree_new, preview_df)

        if warnings:
            self.warn_lbl.config(text="⚠ " + "  ·  ".join(warnings[:3]))
            if not self.warn_lbl.winfo_ismapped():
                self.warn_lbl.pack(fill="x", pady=4, before=self.warn_lbl.master.winfo_children()[1])
        else:
            if self.warn_lbl.winfo_ismapped():
                self.warn_lbl.pack_forget()

    def _fill_tree(self, tree: ttk.Treeview, df: pd.DataFrame):
        tree.delete(*tree.get_children())
        for _, row in df.iterrows():
            vals = []
            for v in row:
                if is_missing(v):
                    vals.append("∅")
                else:
                    s = str(v)
                    vals.append(s if len(s) <= 40 else s[:37] + "…")
            tree.insert("", "end", values=vals)


# ============================================================
#  FENÊTRE PRINCIPALE
# ============================================================

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Pseudonymiseur Excel")
        self.geometry("1280x820")
        self.minsize(960, 600)

        apply_theme(self)

        self.sheets: dict[str, pd.DataFrame] = {}
        self.path: Path | None = None
        self.tabs: dict[str, SheetTab] = {}
        self.existing_map: dict | None = None
        self._preview_after = None

        self._build_toolbar()
        self._build_body()
        self._build_statusbar()
        self._bind_shortcuts()

    # ---- construction UI ----

    def _build_toolbar(self):
        bar = ttk.Frame(self, style="Toolbar.TFrame", padding=(12, 8))
        bar.pack(fill="x")

        # gauche : source
        left = ttk.Frame(bar, style="Toolbar.TFrame")
        left.pack(side="left")

        b1 = ttk.Button(left, text="  Ouvrir Excel", command=self.load_excel)
        b1.pack(side="left", padx=(0, 6))
        Tooltip(b1, "Charger un classeur .xlsx (Ctrl+O)")

        b2 = ttk.Button(left, text="  Réutiliser map…", style="Ghost.TButton",
                        command=self.load_map)
        b2.pack(side="left", padx=(0, 6))
        Tooltip(b2, "Charger une map existante pour préserver les jointures entre fichiers")

        # droite : actions
        right = ttk.Frame(bar, style="Toolbar.TFrame")
        right.pack(side="right")

        b3 = ttk.Button(right, text="Décoder un fichier…",
                        command=self.run_reverse)
        b3.pack(side="right", padx=(6, 0))
        Tooltip(b3, "Restaurer les vraies données à partir d'un fichier traité + sa map")

        b4 = ttk.Button(right, text="🔒  Pseudonymiser", style="Accent.TButton",
                        command=self.run_pseudo)
        b4.pack(side="right")
        Tooltip(b4, "Générer le fichier pseudonymisé + sa map (Ctrl+S)")

        # séparateur
        sep = tk.Frame(self, height=1, bg=THEME["border"])
        sep.pack(fill="x")

    def _build_body(self):
        body = ttk.Frame(self)
        body.pack(fill="both", expand=True)

        # placeholder quand rien n'est chargé
        self.placeholder = ttk.Frame(body)
        self.placeholder.pack(fill="both", expand=True)
        inner = ttk.Frame(self.placeholder)
        inner.place(relx=0.5, rely=0.5, anchor="center")
        ttk.Label(inner, text="Aucun fichier chargé",
                  style="Title.TLabel").pack(pady=(0, 6))
        ttk.Label(inner,
            text="Glissez un fichier Excel ou cliquez sur « Ouvrir Excel »",
            style="Mute.TLabel").pack()
        ttk.Button(inner, text="  Ouvrir Excel  ", style="Accent.TButton",
                   command=self.load_excel).pack(pady=16)

        self.notebook = ttk.Notebook(body)
        # ajouté dynamiquement quand un fichier est chargé

    def _build_statusbar(self):
        sep = tk.Frame(self, height=1, bg=THEME["border"])
        sep.pack(fill="x")
        self.status = ttk.Label(self, text="Prêt", style="Status.TLabel", anchor="w")
        self.status.pack(fill="x")

    def _bind_shortcuts(self):
        self.bind("<Control-o>", lambda e: self.load_excel())
        self.bind("<Control-s>", lambda e: self.run_pseudo())
        self.bind("<F5>", lambda e: self._refresh_active_preview())

    def _set_status(self, text: str, color: str = None):
        self.status.config(text=text, foreground=color or THEME["fg_dim"])

    # ---- preview debounce ----

    def schedule_preview(self):
        if self._preview_after:
            self.after_cancel(self._preview_after)
        self._preview_after = self.after(400, self._refresh_active_preview)

    def _refresh_active_preview(self):
        if not self.tabs: return
        try:
            idx = self.notebook.index("current")
            name = self.notebook.tab(idx, "text")
        except tk.TclError:
            return
        if name in self.tabs:
            self.tabs[name].refresh_preview()

    # ---- actions ----

    def load_excel(self):
        p = filedialog.askopenfilename(
            title="Choisir un classeur Excel",
            filetypes=[("Excel", "*.xlsx *.xlsm")])
        if not p: return
        self.path = Path(p)
        try:
            self.sheets = pd.read_excel(p, sheet_name=None, header=0)
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible de lire :\n{e}")
            return

        # remplace le placeholder par le notebook
        if self.placeholder.winfo_ismapped():
            self.placeholder.pack_forget()
            self.notebook.pack(fill="both", expand=True, padx=8, pady=(8, 0))

        for tab in self.notebook.tabs():
            self.notebook.forget(tab)
        self.tabs.clear()

        for name, df in self.sheets.items():
            tab = SheetTab(self.notebook, self, name, df)
            self.notebook.add(tab, text=f"  {name}  ")
            self.tabs[name] = tab

        nrows = sum(len(df) for df in self.sheets.values())
        self._set_status(
            f"✓  {self.path.name}  ·  {len(self.sheets)} feuille(s)  ·  {nrows} lignes au total",
            THEME["success"])

    def load_map(self):
        p = filedialog.askopenfilename(
            title="Charger une map existante",
            filetypes=[("Map JSON", "*.json")])
        if not p: return
        try:
            self.existing_map = json.loads(Path(p).read_text(encoding="utf-8"))
        except Exception as e:
            messagebox.showerror("Erreur", f"Impossible de lire la map :\n{e}")
            return
        n = len(self.existing_map.get("_shared_lookups", {}))
        self._set_status(f"🔁  Map réutilisée : {n} lookups partagés", THEME["accent"])
        for tab in self.tabs.values(): tab.refresh_preview()

    def run_pseudo(self):
        if not self.sheets:
            messagebox.showwarning("Aucun fichier", "Charge d'abord un classeur Excel.")
            return

        configs = {name: tab.get_config() for name, tab in self.tabs.items()}
        engine = PseudoEngine(self.existing_map)
        try:
            result = engine.apply_workbook(self.sheets, configs)
        except Exception as e:
            messagebox.showerror("Erreur", str(e)); return

        out_path = self.path.with_name(self.path.stem + "_pseudo.xlsx")
        map_path = self.path.with_name(self.path.stem + "_pseudo.map.json")

        try:
            with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
                for name, df in result.items():
                    df.to_excel(writer, sheet_name=name[:31], index=False)
            engine.save(map_path)
        except Exception as e:
            messagebox.showerror("Erreur d'écriture", str(e)); return

        msg = (f"✓ Fichier pseudonymisé :\n   {out_path.name}\n\n"
               f"✓ Map de décodage :\n   {map_path.name}\n\n"
               f"⚠ Garde la map en lieu sûr — sans elle, pas de retour arrière.")
        if engine._warnings:
            msg += "\n\n⚠ Avertissements :\n• " + "\n• ".join(engine._warnings[:10])
            if len(engine._warnings) > 10:
                msg += f"\n…et {len(engine._warnings) - 10} autres."

        messagebox.showinfo("Terminé", msg)
        self._set_status(f"✓  Exporté → {out_path.name}", THEME["success"])

    def run_reverse(self):
        f = filedialog.askopenfilename(
            title="Fichier pseudonymisé (traité par l'IA)",
            filetypes=[("Excel", "*.xlsx")])
        if not f: return
        m = filedialog.askopenfilename(
            title="Map de décodage associée",
            filetypes=[("Map JSON", "*.json")])
        if not m: return

        try:
            sheets_in = pd.read_excel(f, sheet_name=None, header=0)
            engine = PseudoEngine.load(Path(m))
            result = engine.reverse_workbook(sheets_in)
        except Exception as e:
            messagebox.showerror("Erreur", str(e)); return

        out_path = Path(f).with_name(Path(f).stem + "_decoded.xlsx")
        try:
            with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
                for name, df in result.items():
                    df.to_excel(writer, sheet_name=name[:31], index=False)
        except Exception as e:
            messagebox.showerror("Erreur d'écriture", str(e)); return

        msg = f"✓ Fichier décodé :\n   {out_path.name}"
        if engine._warnings:
            msg += "\n\n⚠ " + "\n• ".join(engine._warnings[:5])
        messagebox.showinfo("Terminé", msg)
        self._set_status(f"✓  Décodé → {out_path.name}", THEME["success"])


if __name__ == "__main__":
    App().mainloop()
