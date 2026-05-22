"""
pseudo_excel_streamlit.py — Pseudonymiseur Excel local, multi-feuilles
    pip install streamlit pandas openpyxl
    streamlit run pseudo_excel_streamlit.py
"""
import copy
import json
import re
import secrets
import hashlib
import io

import pandas as pd
import streamlit as st

try:
    from openpyxl import load_workbook
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False


st.set_page_config(layout="wide", page_title="Pseudonymiseur Pro", page_icon="🛡️")


# ============================================================
#  CONSTANTES
# ============================================================

STRATEGIES = {
    "keep":               "Ne rien faire (garder tel quel, formules préservées)",
    "sequential_id":      "Remplacer par ID_1, ID_2…",
    "scale_numeric":      "Multiplier par un facteur (map valeur par valeur)",
    "regex_replace":      "Remplacer une partie via regex",
    "hash_deterministic": "Empreinte stable (hash)",
    "date_shift":         "Décaler toutes les dates (map valeur par valeur)",
    "recalc_formula":     "Recalculer via formule custom",
}

STRATEGY_PARAMS = {
    "keep": [],
    "sequential_id":      [("prefix", "Préfixe", "ID", str)],
    "scale_numeric":      [("factor", "Facteur (vide = aléatoire 0.5–2.0)", "", float)],
    "regex_replace":      [("pattern", "Motif (regex)", r"\d+", str),
                           ("prefix",  "Préfixe",       "N",     str)],
    "hash_deterministic": [("length",  "Longueur (4–64)", 10,   int)],
    "date_shift":         [("days",    "Décalage en jours (vide = aléatoire ±1000)", "", int)],
    "recalc_formula":     [("formula", "Formule (ex: `HT`*(1+`TVA-rate`))", "", str)],
}

DATE_FORMATS_TRY = [
    "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%Y-%m-%d", "%Y/%m/%d",
    "%m/%d/%Y", "%d/%m/%Y %H:%M", "%Y-%m-%d %H:%M:%S",
    "%d %B %Y", "%d %b %Y",
]


# ============================================================
#  HELPERS TYPES / NaN
# ============================================================

def is_missing(v) -> bool:
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
    if pd.api.types.is_numeric_dtype(series):
        return series.astype(float), 0

    def clean(v):
        if is_missing(v):
            return None
        s = str(v).strip()
        s = re.sub(r"[€$£¥\s\u00a0]", "", s)
        if not s:
            return None
        if "," in s and "." in s:
            if s.rfind(",") > s.rfind("."):
                s = s.replace(".", "").replace(",", ".")
            else:
                s = s.replace(",", "")
        elif "," in s:
            s = s.replace(",", ".")
        try:
            return float(s)
        except ValueError:
            return None

    cleaned = series.map(clean)
    n_failed = cleaned.isna().sum() - series.map(is_missing).sum()
    return pd.to_numeric(cleaned, errors="coerce"), int(max(0, n_failed))


def _coerce_param(value, ptype, default=None):
    if value is None: return default
    if isinstance(value, str):
        v = value.strip()
        if v == "": return default
        if ptype is float:
            try: return float(v.replace(",", "."))
            except ValueError: return default
        if ptype is int:
            try: return int(float(v.replace(",", ".")))
            except ValueError: return default
        return v
    try: return ptype(value)
    except (ValueError, TypeError): return default


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    new_cols = []
    seen: dict[str, int] = {}
    for i, c in enumerate(df.columns):
        if c is None or (isinstance(c, float) and pd.isna(c)):
            name = f"_col_{i}"
        elif isinstance(c, str):
            name = c
        else:
            name = str(c)
        if name in seen:
            seen[name] += 1
            name = f"{name}.{seen[name]}"
        else:
            seen[name] = 0
        new_cols.append(name)
    df = df.copy()
    df.columns = new_cols
    return df


def normalize_col_name(name: str) -> str:
    return re.sub(r"\s+", "", str(name)).lower()


# ============================================================
#  CLÉS CANONIQUES POUR MAPS VALUE-BASED
# ============================================================

def canon_number_key(v) -> str | None:
    """Clé canonique pour une valeur numérique. None si non parseable."""
    if is_missing(v): return None
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        try:
            f = float(v)
            if pd.isna(f): return None
            return repr(f)  # repr préserve la précision IEEE 754
        except (ValueError, TypeError):
            return None
    s = str(v).strip()
    if not s: return None
    s = re.sub(r"[€$£¥\s\u00a0]", "", s)
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return repr(float(s))
    except ValueError:
        return None


def canon_date_key(v, hint_format: str | None = None) -> str | None:
    """Clé canonique pour une date (format ISO YYYY-MM-DD HH:MM:SS)."""
    if is_missing(v): return None
    if isinstance(v, pd.Timestamp):
        if pd.isna(v): return None
        return v.strftime("%Y-%m-%d %H:%M:%S")
    if hasattr(v, "strftime") and not isinstance(v, str):
        try: return v.strftime("%Y-%m-%d %H:%M:%S")
        except Exception: return None
    s = str(v).strip()
    if not s: return None
    if hint_format:
        try:
            ts = pd.to_datetime(s, format=hint_format, errors="coerce")
            if pd.notna(ts): return ts.strftime("%Y-%m-%d %H:%M:%S")
        except Exception:
            pass
    try:
        ts = pd.to_datetime(s, errors="coerce", dayfirst=True)
        if pd.notna(ts): return ts.strftime("%Y-%m-%d %H:%M:%S")
    except Exception:
        pass
    return None


# ============================================================
#  LECTURE EXCEL AVEC FORMULES
# ============================================================

def read_excel_with_formulas(file_or_bytes) -> tuple[dict, dict]:
    if hasattr(file_or_bytes, "read"):
        data = file_or_bytes.read()
        file_or_bytes.seek(0)
    else:
        data = file_or_bytes

    sheets_values = pd.read_excel(io.BytesIO(data), sheet_name=None)
    sheets_values = {n: normalize_columns(df) for n, df in sheets_values.items()}

    sheets_formulas: dict[str, dict] = {}
    if not HAS_OPENPYXL:
        return sheets_values, sheets_formulas

    try:
        wb = load_workbook(io.BytesIO(data), data_only=False)
    except Exception:
        return sheets_values, sheets_formulas

    for sheet_name, df in sheets_values.items():
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        formulas = {}
        col_names = list(df.columns)
        for row_idx, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row)):
            for col_idx, cell in enumerate(row):
                if col_idx >= len(col_names): break
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    formulas[(row_idx, col_names[col_idx])] = cell.value
        if formulas:
            sheets_formulas[sheet_name] = formulas
    return sheets_values, sheets_formulas


def write_workbook_with_formulas(sheets: dict, sheets_formulas: dict,
                                  config: dict) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        used_names = set()
        name_mapping = {}
        for name, df in sheets.items():
            safe_name = (str(name)[:31] or "Sheet1")
            base = safe_name
            i = 1
            while safe_name in used_names:
                suffix = f"_{i}"
                safe_name = base[:31 - len(suffix)] + suffix
                i += 1
            used_names.add(safe_name)
            name_mapping[name] = safe_name
            df.to_excel(writer, sheet_name=safe_name, index=False)

        wb = writer.book
        for sheet_name, df in sheets.items():
            if sheet_name not in sheets_formulas:
                continue
            actual_name = name_mapping.get(sheet_name)
            if not actual_name or actual_name not in wb.sheetnames:
                continue
            ws = wb[actual_name]
            sheet_cfg = (config or {}).get(sheet_name, {})
            col_names = list(df.columns)
            col_to_idx = {c: i for i, c in enumerate(col_names)}

            for (row_idx, col_name), formula in sheets_formulas[sheet_name].items():
                strat = sheet_cfg.get(col_name, {}).get("strategy", "keep")
                if strat != "keep":
                    continue
                if col_name not in col_to_idx:
                    continue
                excel_row = row_idx + 2
                excel_col = col_to_idx[col_name] + 1
                ws.cell(row=excel_row, column=excel_col).value = formula

    return buf.getvalue()


# ============================================================
#  HELPERS DATES
# ============================================================

def _try_parse_one(s: str, fmt: str) -> bool:
    try:
        result = pd.to_datetime(s, format=fmt, errors="coerce")
        return pd.notna(result)
    except (ValueError, TypeError):
        return False


def detect_date_format(series: pd.Series) -> str | None:
    samples = []
    for v in series.dropna().head(20):
        if is_missing(v): continue
        if isinstance(v, pd.Timestamp): return None
        if hasattr(v, "strftime") and not isinstance(v, str): return None
        samples.append(str(v))
    if not samples: return None
    best, best_score = None, 0
    for fmt in DATE_FORMATS_TRY:
        score = sum(1 for s in samples if _try_parse_one(s, fmt))
        if score > best_score:
            best, best_score = fmt, score
    return best if best_score >= len(samples) * 0.5 else None


def parse_dates_robust(series: pd.Series, hint_format: str | None = None) -> pd.Series:
    if hint_format:
        try: return pd.to_datetime(series, format=hint_format, errors="coerce")
        except Exception: pass
    fmt = detect_date_format(series)
    dayfirst = bool(fmt and fmt.startswith("%d"))
    try:
        return pd.to_datetime(series, errors="coerce", dayfirst=dayfirst)
    except Exception:
        return pd.Series([pd.NaT] * len(series), index=series.index)


def looks_like_date(v) -> bool:
    """Heuristique : la valeur ressemble-t-elle à une date ?"""
    if is_missing(v): return False
    if isinstance(v, pd.Timestamp): return True
    if hasattr(v, "strftime") and not isinstance(v, str): return True
    if isinstance(v, str):
        s = v.strip()
        if not s or len(s) < 6 or len(s) > 30: return False
        # au moins 2 séparateurs courants OU motif ISO
        return bool(re.search(r"\d{1,4}[/\-.]\d{1,2}[/\-.]\d{1,4}", s))
    return False


# ============================================================
#  HELPERS FORMULES
# ============================================================

def _make_placeholder(placeholders: dict, col: str) -> str:
    for ph, c in placeholders.items():
        if c == col: return ph
    ph = f"__COL_{len(placeholders)}__"
    placeholders[ph] = col
    return ph


def eval_formula(formula: str, df: pd.DataFrame) -> pd.Series:
    placeholders: dict[str, str] = {}
    used_cols: list[str] = []
    expr = formula

    for c in df.columns:
        token = f"`{c}`"
        if token in expr:
            ph = _make_placeholder(placeholders, str(c))
            expr = expr.replace(token, ph)
            used_cols.append(c)

    cols_str = sorted([str(x) for x in df.columns], key=len, reverse=True)
    for c in cols_str:
        if any(c == cx for cx in placeholders.values()): continue
        pattern = rf"(?<![A-Za-z0-9_]){re.escape(c)}(?![A-Za-z0-9_])"
        new_expr, n = re.subn(
            pattern,
            lambda m, col=c: _make_placeholder(placeholders, col),
            expr,
        )
        if n > 0:
            expr = new_expr
            used_cols.append(c)

    for ph, col in placeholders.items():
        expr = expr.replace(ph, f"__df__[{col!r}]")

    safe_df = df.copy()
    for c in set(used_cols):
        if c in safe_df.columns and not pd.api.types.is_numeric_dtype(safe_df[c]):
            safe_df[c], _ = coerce_numeric_smart(safe_df[c])

    try:
        return eval(expr, {"__builtins__": {}}, {"__df__": safe_df})
    except KeyError as e:
        raise ValueError(f"Colonne inexistante référencée par la formule : {e}")
    except Exception as e:
        raise ValueError(f"Erreur dans la formule '{formula}' : {e}")


# ============================================================
#  MOTEUR — 100% value-based pour le reverse
# ============================================================

class PseudoEngine:
    """
    Map :
      {
        "_shared_lookups":   { "<col_canon>": {src: pseudo}, ... },     # strings
        "_number_map":       { "<repr(orig)>": pseudo_num, ... },       # GLOBAL value-based
        "_date_map":         { "<iso_orig>": "<iso_pseudo>", ... },     # GLOBAL value-based
        "_scale_factors":    { "<col_canon>": 1.234 },                  # fallback only
        "_date_shifts":      { "<col_canon>": 42 },                     # fallback only
        "_date_formats":     { "<col_canon>": "%d/%m/%Y" },             # pour réécriture
        "sheets":            { "<sheet>": { "<col>": {"strategy":...} } }
      }
    """

    def __init__(self, existing: dict | None = None):
        self.map = existing if existing else {}
        self.map.setdefault("_shared_lookups", {})
        self.map.setdefault("_number_map", {})
        self.map.setdefault("_date_map", {})
        self.map.setdefault("_scale_factors", {})
        self.map.setdefault("_date_shifts", {})
        self.map.setdefault("_date_formats", {})
        self.map.setdefault("sheets", {})
        self._migrate_legacy_map()
        self._warnings: list[str] = []
        # Caches d'inverses pour le reverse
        self._inv_numbers: dict | None = None
        self._inv_dates: dict | None = None
        self._inv_strings_global: dict | None = None

    def _migrate_legacy_map(self):
        sl = self.map.get("_shared_lookups", {})
        moved = []
        for k, v in list(sl.items()):
            if k.startswith("__factor__") and isinstance(v, dict) and "value" in v:
                col = k[len("__factor__"):]
                self.map["_scale_factors"][normalize_col_name(col)] = float(v["value"])
                moved.append(k)
            elif k.startswith("__date_shift__") and isinstance(v, dict) and "days" in v:
                col = k[len("__date_shift__"):]
                self.map["_date_shifts"][normalize_col_name(col)] = int(v["days"])
                moved.append(k)
        for k in moved:
            del sl[k]

    def _lookup_for(self, col: str) -> dict:
        return self.map["_shared_lookups"].setdefault(normalize_col_name(col), {})

    def _build_inverses(self):
        if self._inv_numbers is None:
            # Inverse des nombres : pseudo_repr → repr(orig)
            self._inv_numbers = {}
            for orig_repr, pseudo in self.map["_number_map"].items():
                pseudo_repr = canon_number_key(pseudo)
                if pseudo_repr is not None:
                    self._inv_numbers[pseudo_repr] = orig_repr
        if self._inv_dates is None:
            # Inverse des dates : iso_pseudo → iso_orig
            self._inv_dates = {v: k for k, v in self.map["_date_map"].items()}
        if self._inv_strings_global is None:
            # Inverse des strings : pseudo → orig (tous lookups confondus)
            self._inv_strings_global = {}
            for col_lu in self.map["_shared_lookups"].values():
                for src, pseudo in col_lu.items():
                    self._inv_strings_global[pseudo] = src

    # ---- pipeline apply ----

    def apply_sheet(self, df: pd.DataFrame, sheet: str, config: dict) -> pd.DataFrame:
        out = df.copy()
        sheet_map = self.map["sheets"].setdefault(sheet, {})

        for col, cfg in config.items():
            if col not in out.columns: continue
            strat = cfg.get("strategy", "keep")
            p = cfg.get("params", {}) or {}
            sheet_map[col] = {"strategy": strat, "params": p}
            canon = normalize_col_name(col)

            if strat in ("keep", "recalc_formula"):
                continue

            elif strat == "sequential_id":
                prefix = p.get("prefix") or "ID"
                lu = self._lookup_for(col)
                out[col] = pd.Series(
                    [(None if is_missing(v)
                      else lu.setdefault(str(v).strip(), f"{prefix}_{len(lu)+1}"))
                     for v in out[col]],
                    index=out.index, dtype="object",
                )

            elif strat == "scale_numeric":
                factor = _coerce_param(p.get("factor"), float)
                if factor is None or factor == 0:
                    factor = round(0.5 + secrets.randbelow(1500) / 1000, 4)
                self.map["_scale_factors"].setdefault(canon, factor)
                used_factor = self.map["_scale_factors"][canon]
                numeric, n_failed = coerce_numeric_smart(out[col])
                if n_failed > 0:
                    self._warnings.append(
                        f"Colonne '{col}' ({sheet}) : {n_failed} valeur(s) "
                        f"non numérique(s) conservée(s) telles quelles."
                    )
                res = []
                num_map = self.map["_number_map"]
                for orig_v, num_v in zip(out[col], numeric):
                    if is_missing(orig_v):
                        res.append(None); continue
                    if pd.isna(num_v):
                        res.append(orig_v); continue
                    orig_key = canon_number_key(orig_v)
                    if orig_key is None:
                        res.append(orig_v); continue
                    # Stocke la map valeur par valeur
                    if orig_key in num_map:
                        new_val = num_map[orig_key]
                    else:
                        new_val = float(num_v) * used_factor
                        num_map[orig_key] = new_val
                    res.append(new_val)
                out[col] = pd.Series(res, index=out.index, dtype="object")

            elif strat == "regex_replace":
                pat = p.get("pattern") or r"\d+"
                pref = p.get("prefix") or "N"
                try:
                    rx = re.compile(pat)
                except re.error as e:
                    self._warnings.append(
                        f"Colonne '{col}' ({sheet}) : regex invalide ({e}), ignorée."
                    )
                    continue
                lu = self._lookup_for(col)
                out[col] = out[col].map(
                    lambda v: None if is_missing(v)
                    else rx.sub(
                        lambda m: lu.setdefault(m.group(0), f"{pref}_{len(lu)+1}"),
                        str(v),
                    )
                )

            elif strat == "hash_deterministic":
                length = _coerce_param(p.get("length"), int, default=10)
                if length is None or length < 4: length = 10
                if length > 64: length = 64
                lu = self._lookup_for(col)
                out[col] = pd.Series(
                    [(None if is_missing(v)
                      else lu.setdefault(
                          str(v).strip(),
                          hashlib.sha256(str(v).strip().encode("utf-8")).hexdigest()[:length],
                      ))
                     for v in out[col]],
                    index=out.index, dtype="object",
                )

            elif strat == "date_shift":
                days = _coerce_param(p.get("days"), int)
                if days is None:
                    days = secrets.randbelow(2000) - 1000
                self.map["_date_shifts"].setdefault(canon, days)
                used_days = self.map["_date_shifts"][canon]
                fmt = detect_date_format(out[col])
                self.map["_date_formats"][canon] = fmt
                parsed = parse_dates_robust(out[col], hint_format=fmt)

                non_missing = (~out[col].map(is_missing)).sum()
                n_parsed = parsed.notna().sum()
                if non_missing > 0 and n_parsed == 0:
                    self._warnings.append(
                        f"Colonne '{col}' ({sheet}) : aucune date reconnue, ignorée."
                    )
                    continue
                n_failed = int(non_missing - n_parsed)
                if n_failed > 0:
                    self._warnings.append(
                        f"Colonne '{col}' ({sheet}) : {n_failed} date(s) non reconnue(s)."
                    )

                shifted = parsed + pd.Timedelta(days=used_days)
                date_map = self.map["_date_map"]
                res = []
                for orig_v, sh in zip(out[col], shifted):
                    if is_missing(orig_v):
                        res.append(None); continue
                    if pd.isna(sh):
                        res.append(orig_v); continue
                    orig_key = canon_date_key(orig_v, hint_format=fmt)
                    pseudo_key = sh.strftime("%Y-%m-%d %H:%M:%S")
                    if orig_key:
                        date_map[orig_key] = pseudo_key
                    res.append(sh.strftime(fmt) if fmt else sh)
                out[col] = pd.Series(res, index=out.index, dtype="object")

        # Recalculs
        for col, cfg in config.items():
            if col not in out.columns: continue
            if cfg.get("strategy") == "recalc_formula":
                formula = (cfg.get("params", {}) or {}).get("formula", "")
                if formula and formula.strip():
                    try:
                        out[col] = eval_formula(formula, out)
                    except ValueError as e:
                        self._warnings.append(f"Colonne '{col}' ({sheet}) : {e}")

        # Invalide les caches d'inverses (la map a changé)
        self._inv_numbers = self._inv_dates = self._inv_strings_global = None
        return out

    def apply_workbook(self, sheets: dict, configs: dict) -> dict:
        return {n: self.apply_sheet(df, n, configs.get(n, {})) for n, df in sheets.items()}

    # ---- reverse 100% value-based ----

    def reverse_value(self, v):
        """
        Tente de retrouver la valeur originale d'une cellule.
        Ordre : string global > number map > date map > inchangé.
        Retourne (nouvelle_valeur, a_changé).
        """
        if is_missing(v):
            return None, False

        # 1. String simple (clé directe)
        if isinstance(v, str) and v in self._inv_strings_global:
            return self._inv_strings_global[v], True

        # 2. Number map
        num_key = canon_number_key(v)
        if num_key is not None and num_key in self._inv_numbers:
            orig_repr = self._inv_numbers[num_key]
            try:
                return float(orig_repr), True
            except ValueError:
                pass

        # 3. Date map
        if looks_like_date(v):
            date_key = canon_date_key(v)
            if date_key and date_key in self._inv_dates:
                orig_iso = self._inv_dates[date_key]
                # On essaye de retrouver le format d'origine de la première colonne match
                # Sinon, format ISO par défaut
                try:
                    ts = pd.to_datetime(orig_iso)
                    return ts, True
                except Exception:
                    return orig_iso, True

        # 4. String composite (regex sub) — pour les libellés type "VIR-N_42"
        if isinstance(v, str) and self._inv_strings_global:
            pseudos = sorted(self._inv_strings_global.keys(), key=len, reverse=True)
            if pseudos:
                rx = re.compile("|".join(re.escape(p) for p in pseudos))
                new_s = rx.sub(lambda m: self._inv_strings_global[m.group(0)], v)
                if new_s != v:
                    return new_s, True

        return v, False

    def reverse_sheet(self, df: pd.DataFrame, sheet: str | None = None) -> pd.DataFrame:
        """Décodage purement value-based, fonctionne sur n'importe quelle feuille."""
        self._build_inverses()
        out = df.copy()

        for col in out.columns:
            canon = normalize_col_name(col)
            new_col = []
            n_changed = 0
            for v in out[col]:
                nv, changed = self.reverse_value(v)
                if changed: n_changed += 1
                new_col.append(nv)
            out[col] = pd.Series(new_col, index=out.index, dtype="object")

            # Reformatage des dates si on connaît un format pour cette colonne
            fmt = self.map["_date_formats"].get(canon)
            if fmt:
                def reformat(v):
                    if isinstance(v, pd.Timestamp): return v.strftime(fmt)
                    return v
                out[col] = out[col].map(reformat)

            # Fallback scale_numeric : si la colonne porte un nom connu et que peu
            # de valeurs ont été retrouvées par la map, on tente le facteur.
            if (n_changed == 0 and canon in self.map["_scale_factors"]):
                factor = float(self.map["_scale_factors"][canon])
                if factor != 0:
                    num, _ = coerce_numeric_smart(out[col])
                    if num.notna().any():
                        res = out[col].copy().astype(object)
                        mask = num.notna()
                        res[mask] = num[mask] / factor
                        out[col] = res
                        self._warnings.append(
                            f"Colonne '{col}' : valeurs absentes de la map, "
                            f"décodage par facteur d'échelle (best-effort)."
                        )

        return out

    def reverse_workbook(self, sheets: dict) -> dict:
        return {n: self.reverse_sheet(df, n) for n, df in sheets.items()}


# ============================================================
#  HELPERS UI
# ============================================================

def init_session():
    defaults = {
        "df_dict": None,
        "formulas_dict": None,
        "config": None,
        "map_data": None,
        "sheet_names": None,
        "file_name": None,
        "mode": "pseudonymiser",
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


def suggest_strategy(col: str) -> str:
    cl = str(col).lower()
    tokens = re.split(r"[\s_\-./]+", cl)
    has = lambda *kws: any(t in kws for t in tokens)
    if "ttc" in tokens: return "recalc_formula"
    if has("tva", "rate", "taux", "vat"): return "keep"
    if has("ht", "montant", "prix", "amount", "price", "total"): return "scale_numeric"
    if has("id", "client", "nom", "name", "customer"): return "sequential_id"
    if has("date"): return "date_shift"
    if has("libelle", "libellé", "ref", "label", "reference"): return "regex_replace"
    return "keep"


def make_default_config(df_dict: dict, formulas_dict: dict) -> dict:
    cfg = {}
    for sheet, df in df_dict.items():
        cfg[sheet] = {}
        cols_with_formula = set()
        for (_, col_name) in formulas_dict.get(sheet, {}).keys():
            cols_with_formula.add(col_name)
        for c in df.columns:
            if c in cols_with_formula:
                cfg[sheet][c] = {"strategy": "keep", "params": {}}
            else:
                cfg[sheet][c] = {"strategy": suggest_strategy(c), "params": {}}
    return cfg


def smart_default(col: str, strat: str, all_cols: list, pname: str):
    cl = str(col).lower()
    tokens = re.split(r"[\s_\-./]+", cl)
    has = lambda *kws: any(t in kws for t in tokens)
    if strat == "sequential_id" and pname == "prefix":
        if has("client"): return "CLI"
        if has("nom", "name"): return "NOM"
        if has("id"): return "ID"
        return str(col).upper()[:5] if col else "ID"
    if strat == "recalc_formula" and pname == "formula" and "ttc" in tokens:
        ht = next((c for c in all_cols
                   if "ht" in re.split(r"[\s_\-./]+", str(c).lower())), "HT")
        tva = next((c for c in all_cols
                    if any(k in re.split(r"[\s_\-./]+", str(c).lower())
                           for k in ("tva", "rate", "taux", "vat"))), "TVA-rate")
        return f"`{ht}`*(1+`{tva}`)"
    if strat == "regex_replace":
        if pname == "pattern": return r"\d+"
        if pname == "prefix": return "N"
    return None


def merge_map_into_config(map_data: dict, config: dict) -> dict:
    if not map_data or not config: return config
    for sheet, sheet_cfg in map_data.get("sheets", {}).items():
        if sheet not in config: continue
        for col, entry in sheet_cfg.items():
            if col in config[sheet]:
                strat = entry.get("strategy", "keep")
                if strat in STRATEGIES:
                    config[sheet][col] = {
                        "strategy": strat,
                        "params": entry.get("params", {}) or {},
                    }
    return config


def safe_dataframe_for_display(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    for c in out.columns:
        if out[c].dtype == "object":
            out[c] = out[c].map(lambda v: "" if is_missing(v) else str(v))
    return out


def has_formula_in_column(formulas_sheet: dict, col_name: str) -> int:
    return sum(1 for (_, c) in formulas_sheet.keys() if c == col_name)


# ============================================================
#  APP
# ============================================================

init_session()

st.title("🛡️ Pseudonymiseur Excel — local")
st.caption("Pseudonymisation **value-based** : la map stocke chaque correspondance, "
           "donc le décodage marche même sur des feuilles ajoutées par l'IA.")

mode = st.radio(
    "Mode",
    ["pseudonymiser", "decoder"],
    horizontal=True,
    format_func=lambda m: "🔒 Pseudonymiser" if m == "pseudonymiser" else "🔓 Décoder",
    key="mode",
)
st.divider()


# ============================================================
#  MODE DÉCODAGE
# ============================================================

if mode == "decoder":
    st.subheader("🔓 Décoder un fichier (potentiellement transformé par l'IA)")
    st.caption("Toutes les feuilles du fichier sont décodées, **y compris celles "
               "ajoutées par l'IA** (synthèses, pivots, etc.). Le décodage cherche "
               "chaque valeur dans la map, indépendamment du nom de feuille ou de colonne.")
    col1, col2 = st.columns(2)
    with col1:
        dec_excel = st.file_uploader("Fichier Excel", type=["xlsx"], key="dec_xl")
    with col2:
        dec_map = st.file_uploader("Map JSON associée", type=["json"], key="dec_map")

    if dec_excel and dec_map:
        try:
            sheets_in = pd.read_excel(dec_excel, sheet_name=None)
            sheets_in = {name: normalize_columns(df) for name, df in sheets_in.items()}
            map_data = json.load(dec_map)
            engine = PseudoEngine(map_data)
            result = engine.reverse_workbook(sheets_in)

            st.success(f"✓ Décodage effectué sur {len(result)} feuille(s).")
            if engine._warnings:
                with st.expander(f"⚠ {len(engine._warnings)} avertissement(s)"):
                    for w in engine._warnings:
                        st.warning(w)

            for name, df in result.items():
                with st.expander(f"📄 {name} ({len(df)} lignes)", expanded=False):
                    st.dataframe(safe_dataframe_for_display(df.head(50)),
                                 use_container_width=True, hide_index=True)

            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as w:
                used = set()
                for n, df in result.items():
                    sn = (str(n)[:31] or "Sheet1")
                    base = sn; i = 1
                    while sn in used:
                        sn = base[:31 - len(f"_{i}")] + f"_{i}"; i += 1
                    used.add(sn)
                    df.to_excel(w, sheet_name=sn, index=False)
            st.download_button(
                "📥 Télécharger Excel décodé",
                data=buf.getvalue(),
                file_name="decoded_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        except Exception as e:
            st.error(f"Erreur : {e}")
    else:
        st.info("Fournis le fichier Excel **et** sa map JSON.")
    st.stop()


# ============================================================
#  MODE PSEUDONYMISATION
# ============================================================

with st.container():
    col_up1, col_up2 = st.columns(2)
    with col_up1:
        uploaded_excel = st.file_uploader("Fichier Excel à pseudonymiser (.xlsx)", type=["xlsx"])
    with col_up2:
        uploaded_map = st.file_uploader("Map existante à réutiliser (optionnel)", type=["json"])

if uploaded_excel and st.session_state.file_name != uploaded_excel.name:
    try:
        df_dict, formulas_dict = read_excel_with_formulas(uploaded_excel)
    except Exception as e:
        st.error(f"Impossible de lire le fichier : {e}")
        st.stop()
    st.session_state.df_dict = df_dict
    st.session_state.formulas_dict = formulas_dict
    st.session_state.sheet_names = list(df_dict.keys())
    st.session_state.file_name = uploaded_excel.name
    st.session_state.config = make_default_config(df_dict, formulas_dict)
    for k in list(st.session_state.keys()):
        if (k.startswith("col_page_") or k.startswith("col_offset_")
            or k.startswith("filter_") or k.startswith("row_range_")
            or k.startswith("strat_") or k.startswith("param_")
            or k.startswith("offsetslider_") or k.startswith("prev_")
            or k.startswith("next_")):
            del st.session_state[k]
    for k in ("final_excel", "final_map", "final_warnings"):
        if k in st.session_state: del st.session_state[k]
    st.rerun()

if uploaded_map is not None:
    try:
        loaded = json.load(uploaded_map)
        st.session_state.map_data = loaded
        if st.session_state.config:
            st.session_state.config = merge_map_into_config(loaded, st.session_state.config)
    except Exception as e:
        st.error(f"Map JSON invalide : {e}")

if not st.session_state.df_dict:
    st.info("Charge un fichier Excel pour commencer.")
    st.stop()

st.divider()

sheet = st.sidebar.selectbox("📂 Feuille active", st.session_state.sheet_names)
df_current = st.session_state.df_dict[sheet]
formulas_current = (st.session_state.formulas_dict or {}).get(sheet, {})

n_formulas_total = len(formulas_current)
if n_formulas_total > 0:
    st.sidebar.info(f"📐 {n_formulas_total} formule(s) Excel détectée(s). "
                    "Préservées sur les colonnes 'keep'.")

col_gauche, col_droite = st.columns([1, 2], gap="large")


# ─── COLONNE GAUCHE ───
with col_gauche:
    st.subheader("🚀 Export")

    if st.button("🔒 Pseudonymiser tout le classeur",
                 use_container_width=True, type="primary"):
        with st.spinner("Traitement..."):
            try:
                final_engine = PseudoEngine(copy.deepcopy(st.session_state.map_data))
                processed = final_engine.apply_workbook(
                    st.session_state.df_dict,
                    st.session_state.config,
                )
                excel_bytes = write_workbook_with_formulas(
                    processed,
                    st.session_state.formulas_dict or {},
                    st.session_state.config,
                )
                map_bytes = json.dumps(
                    final_engine.map, indent=2, ensure_ascii=False, default=str
                ).encode("utf-8")
                st.session_state.final_excel = excel_bytes
                st.session_state.final_map = map_bytes
                st.session_state.final_warnings = list(final_engine._warnings)
                st.success("✓ Fichier généré.")
            except Exception as e:
                st.error(f"Erreur : {e}")

    if "final_excel" in st.session_state:
        d1, d2 = st.columns(2)
        with d1:
            st.download_button(
                "📥 Excel pseudo",
                data=st.session_state.final_excel,
                file_name="pseudonymise_output.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with d2:
            st.download_button(
                "📥 Map JSON",
                data=st.session_state.final_map,
                file_name="map_pseudo.json",
                mime="application/json",
                use_container_width=True,
            )
        if st.session_state.get("final_warnings"):
            with st.expander(f"⚠ {len(st.session_state.final_warnings)} avertissement(s)"):
                for w in st.session_state.final_warnings:
                    st.warning(w)

    st.divider()
    st.subheader("⚙️ Stratégies")
    st.caption("Modifications appliquées en direct.")

    scroll_config = st.container(height=650)
    with scroll_config:
        all_cols = list(df_current.columns)
        for c_name in all_cols:
            if not isinstance(c_name, str): continue
            if c_name not in df_current.columns: continue

            cdict = st.session_state.config[sheet].get(
                c_name, {"strategy": "keep", "params": {}}
            )
            current_strat = cdict.get("strategy", "keep")
            if current_strat not in STRATEGIES: current_strat = "keep"

            n_form = has_formula_in_column(formulas_current, c_name)
            label_prefix = "📐 " if n_form > 0 else "🔠 "
            label_suffix = f" — {STRATEGIES[current_strat]}"
            if n_form > 0:
                label_suffix += f" · {n_form} formule(s)"

            with st.expander(f"{label_prefix}{c_name}{label_suffix}", expanded=False):
                if n_form > 0:
                    st.info(f"Cette colonne contient {n_form} formule(s) Excel. "
                            f"En 'keep', elles seront restaurées dans l'export.")
                    samples_in_col = [(r, f) for (r, c), f in formulas_current.items()
                                      if c == c_name][:3]
                    for r, f in samples_in_col:
                        st.code(f"L{r+2}: {f}", language="text")

                new_strat = st.selectbox(
                    "Stratégie",
                    options=list(STRATEGIES.keys()),
                    format_func=lambda x: STRATEGIES[x],
                    index=list(STRATEGIES.keys()).index(current_strat),
                    key=f"strat_{sheet}_{c_name}",
                )

                try:
                    series = df_current[c_name]
                    samples = [str(v) for v in series.dropna().head(3) if not is_missing(v)]
                    if samples:
                        st.caption("Exemples : " + " · ".join(s[:20] for s in samples))
                except Exception:
                    st.caption("(aperçu indisponible)")

                new_params = {}
                for p_name, p_label, p_default, p_type in STRATEGY_PARAMS[new_strat]:
                    smart = smart_default(c_name, new_strat, all_cols, p_name)
                    fallback = smart if smart is not None else p_default
                    stored = cdict.get("params", {}).get(p_name)
                    if stored in (None, ""):
                        initial = str(fallback) if fallback != "" else ""
                    else:
                        initial = str(stored)
                    val = st.text_input(
                        p_label, value=initial,
                        key=f"param_{sheet}_{c_name}_{p_name}_{new_strat}",
                    )
                    new_params[p_name] = val

                st.session_state.config[sheet][c_name] = {
                    "strategy": new_strat, "params": new_params,
                }


# ─── COLONNE DROITE ───
with col_droite:
    st.subheader("👁️ Aperçu en direct")

    with st.expander("🔍 Filtres (vide = tout afficher)"):
        filter_cols = st.multiselect(
            "Colonnes à filtrer",
            list(df_current.columns),
            key=f"filter_cols_{sheet}",
        )
        df_filtered = df_current.copy()
        for fcol in filter_cols:
            if fcol not in df_filtered.columns: continue
            try:
                unique_vals = sorted({
                    str(v) for v in df_filtered[fcol].dropna() if not is_missing(v)
                })
            except Exception:
                unique_vals = []
            selected = st.multiselect(
                f"Valeurs pour '{fcol}'", unique_vals, default=[],
                key=f"filter_{sheet}_{fcol}",
            )
            if selected:
                df_filtered = df_filtered[
                    df_filtered[fcol].astype(str).isin(selected)
                ]

    if len(df_filtered) == 0:
        st.warning("Aucune ligne ne correspond aux filtres.")
        st.stop()

    max_rows = len(df_filtered)
    if max_rows > 1:
        row_range = st.select_slider(
            "Plage de lignes",
            options=list(range(max_rows + 1)),
            value=(0, min(100, max_rows)),
            key=f"row_range_{sheet}",
        )
        df_snippet = df_filtered.iloc[row_range[0]:row_range[1]].copy()
    else:
        df_snippet = df_filtered.copy()

    try:
        engine = PseudoEngine(copy.deepcopy(st.session_state.map_data))
        df_preview_full = engine.apply_sheet(
            df_snippet, sheet, st.session_state.config[sheet]
        )
    except Exception as e:
        st.error(f"Erreur d'aperçu : {e}")
        st.stop()

    # Pagination flèches + slider indépendants
    max_cols = len(df_preview_full.columns)
    cols_per_page = 5
    page_key = f"col_page_{sheet}"
    offset_key = f"col_offset_{sheet}"
    if page_key not in st.session_state: st.session_state[page_key] = 0
    if offset_key not in st.session_state: st.session_state[offset_key] = 0

    n_pages = max(1, (max_cols + cols_per_page - 1) // cols_per_page)
    if st.session_state[page_key] >= n_pages:
        st.session_state[page_key] = 0

    if max_cols > cols_per_page:
        nav1, nav2, nav3 = st.columns([1, 3, 1], vertical_alignment="center")
        with nav1:
            if st.button("⬅️ Page", disabled=st.session_state[page_key] <= 0,
                         use_container_width=True, key=f"prev_{sheet}"):
                st.session_state[page_key] -= 1
                st.rerun()
        with nav2:
            max_offset = max(0, min(cols_per_page - 1, max_cols - 1))
            if max_offset > 0:
                new_offset = st.slider(
                    "Décalage fin", 0, max_offset,
                    min(st.session_state[offset_key], max_offset),
                    label_visibility="collapsed",
                    key=f"offsetslider_{sheet}",
                )
                st.session_state[offset_key] = new_offset
            else:
                st.session_state[offset_key] = 0
            page = st.session_state[page_key]
            offset = st.session_state[offset_key]
            start = min(page * cols_per_page + offset, max(0, max_cols - 1))
            end = min(start + cols_per_page, max_cols)
            st.markdown(
                f"<div style='text-align:center;color:#888'>"
                f"Page <b>{page+1}/{n_pages}</b> · décalage <b>+{offset}</b> · "
                f"colonnes <b>{start+1}–{end}</b> sur {max_cols}"
                f"</div>", unsafe_allow_html=True,
            )
        with nav3:
            if st.button("Page ➡️",
                         disabled=st.session_state[page_key] >= n_pages - 1,
                         use_container_width=True, key=f"next_{sheet}"):
                st.session_state[page_key] += 1
                st.rerun()
    else:
        st.session_state[page_key] = 0
        st.session_state[offset_key] = 0

    page = st.session_state[page_key]
    offset = st.session_state[offset_key]
    cs = min(page * cols_per_page + offset, max(0, max_cols - 1))
    df_preview_view = df_preview_full.iloc[:, cs:cs + cols_per_page]
    df_snippet_view = df_snippet.iloc[:, cs:cs + cols_per_page]

    tab_apres, tab_avant, tab_form = st.tabs([
        "✨ Après pseudonymisation",
        "📋 Original",
        f"📐 Formules ({n_formulas_total})",
    ])
    with tab_apres:
        if engine._warnings:
            for w in engine._warnings[:5]: st.warning(w)
            if len(engine._warnings) > 5:
                st.caption(f"…et {len(engine._warnings)-5} autres.")
        st.dataframe(safe_dataframe_for_display(df_preview_view),
                     use_container_width=True, height=600, hide_index=True)
    with tab_avant:
        st.dataframe(safe_dataframe_for_display(df_snippet_view),
                     use_container_width=True, height=600, hide_index=True)
    with tab_form:
        if not formulas_current:
            st.info("Aucune formule détectée dans cette feuille.")
        else:
            st.caption(f"{n_formulas_total} formule(s). Préservées si stratégie 'keep'.")
            df_form = pd.DataFrame([
                {"Ligne Excel": r + 2, "Colonne": c, "Formule": f,
                 "Stratégie": st.session_state.config[sheet].get(c, {}).get("strategy", "keep")}
                for (r, c), f in formulas_current.items()
            ]).sort_values(["Ligne Excel", "Colonne"])
            st.dataframe(df_form, use_container_width=True, hide_index=True, height=600)

    st.caption(
        "ℹ️ Aperçu = pseudos calculés sur la plage visible. "
        "L'export utilise tout le classeur (mêmes valeurs → mêmes pseudos)."
    )
