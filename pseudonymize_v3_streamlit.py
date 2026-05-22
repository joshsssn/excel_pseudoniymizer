import copy
import json
import re
import secrets
import hashlib
from pathlib import Path
import io

import pandas as pd
import streamlit as st


# ============================================================
#  CONFIGURATION STREAMLIT
# ============================================================
st.set_page_config(layout="wide", page_title="Pseudonymiseur Pro", page_icon="🛡️")

# ============================================================
#  MOTEUR & LOGIQUE METIER
# ============================================================

STRATEGIES = {
    "keep":               "Ne rien faire (garder tel quel)",
    "sequential_id":      "Remplacer par ID_1, ID_2…",
    "scale_numeric":      "Multiplier par un facteur",
    "regex_replace":      "Remplacer une partie via regex",
    "hash_deterministic": "Empreinte stable (hash)",
    "date_shift":         "Décaler toutes les dates",
    "recalc_formula":     "Recalculer via formule",
}

STRATEGY_PARAMS = {
    "keep": [],
    "sequential_id": [("prefix", "Préfixe", "ID", str)],
    "scale_numeric": [("factor", "Facteur (vide=aléatoire)", "", float)],
    "regex_replace": [("pattern", "Motif (regex)", r"\d+", str), ("prefix", "Préfixe", "N", str)],
    "hash_deterministic": [("length", "Longueur", 10, int)],
    "date_shift": [("days", "Décalage jours (vide=aléatoire)", "", int)],
    "recalc_formula": [("formula", "Formule", "", str)],
}

DATE_FORMATS_TRY = [
    "%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%Y-%m-%d", "%Y/%m/%d",
    "%m/%d/%Y", "%d/%m/%Y %H:%M", "%Y-%m-%d %H:%M:%S", "%d %B %Y", "%d %b %Y",
]

def is_missing(v) -> bool:
    if v is None: return True
    try:
        if pd.isna(v): return True
    except (TypeError, ValueError): pass
    if isinstance(v, str) and v.strip().lower() in ("", "nan", "nat", "none", "<na>", "null"):
        return True
    return False

def coerce_numeric_smart(series: pd.Series) -> tuple[pd.Series, int]:
    if pd.api.types.is_numeric_dtype(series):
        return series.astype(float), 0
    def clean(v):
        if is_missing(v): return None
        s = str(v).strip()
        s = re.sub(r"[€$£¥\s\u00a0]", "", s)
        if not s: return None
        if "," in s and "." in s:
            if s.rfind(",") > s.rfind("."): s = s.replace(".", "").replace(",", ".")
            else: s = s.replace(",", "")
        elif "," in s: s = s.replace(",", ".")
        try: return float(s)
        except ValueError: return None
    cleaned = series.map(clean)
    n_failed = cleaned.isna().sum() - series.map(is_missing).sum()
    return pd.to_numeric(cleaned, errors="coerce"), int(max(0, n_failed))

def detect_date_format(series: pd.Series) -> str | None:
    samples = [str(v) for v in series.dropna().head(20) if not is_missing(v)]
    if not samples: return None
    best, best_score = None, 0
    for fmt in DATE_FORMATS_TRY:
        score = sum(1 for s in samples if (lambda x: pd.to_datetime(x, format=fmt, errors="ignore"))(s) is not s)
        if score > best_score:
            best, best_score = fmt, score
    return best if best_score >= len(samples) * 0.5 else None

def parse_dates_robust(series: pd.Series, hint_format: str | None = None) -> pd.Series:
    if hint_format:
        try: return pd.to_datetime(series, format=hint_format, errors="coerce")
        except Exception: pass
    fmt = detect_date_format(series)
    dayfirst = bool(fmt and fmt.startswith("%d"))
    try: return pd.to_datetime(series, errors="coerce", dayfirst=dayfirst)
    except Exception: return pd.Series([pd.NaT] * len(series), index=series.index)

def eval_formula(formula: str, df: pd.DataFrame) -> pd.Series:
    placeholders, used_cols = {}, []
    expr = formula
    for c in df.columns:
        token = f"`{c}`"
        if token in expr:
            ph = f"__COL_{len(placeholders)}__"
            placeholders[ph] = str(c)
            expr = expr.replace(token, ph)
            used_cols.append(c)
    for c in sorted([str(x) for x in df.columns], key=len, reverse=True):
        if any(p == c for p in placeholders.values()): continue
        pattern = rf"(?<![A-Za-z0-9_]){re.escape(c)}(?![A-Za-z0-9_])"
        new_expr, n = re.subn(pattern, lambda m, col=c: (lambda p, c: next((ph for ph, cx in p.items() if cx == c), (p.setdefault(f"__COL_{len(p)}__", c) and f"__COL_{len(p)-1}__")))(placeholders, col), expr)
        if n > 0:
            expr = new_expr
            used_cols.append(c)
    for ph, col in placeholders.items():
        expr = expr.replace(ph, f"__df__[{col!r}]")
    safe_df = df.copy()
    for c in set(used_cols):
        if c in safe_df.columns and not pd.api.types.is_numeric_dtype(safe_df[c]):
            safe_df[c], _ = coerce_numeric_smart(safe_df[c])
    try: return eval(expr, {"__builtins__": {}}, {"__df__": safe_df})
    except Exception as e: raise ValueError(f"Erreur formule '{formula}': {e}")

class PseudoEngine:
    def __init__(self, existing: dict | None = None):
        self.map = existing if existing else {"_shared_lookups": {}, "_date_formats": {}, "sheets": {}}
        self.map.setdefault("_shared_lookups", {})
        self.map.setdefault("_date_formats", {})
        self.map.setdefault("sheets", {})
        self._warnings = []

    def _get_lookup(self, col: str): return self.map["_shared_lookups"].setdefault(col, {})

    def apply_sheet(self, df: pd.DataFrame, sheet: str, config: dict) -> pd.DataFrame:
        out = df.copy()
        sheet_map = self.map["sheets"].setdefault(sheet, {})
        for col, cfg in config.items():
            strat, p = cfg["strategy"], cfg.get("params", {})
            sheet_map[col] = {"strategy": strat, "params": p, "lookup_key": col}
            if strat == "sequential_id":
                prefix = p.get("prefix", "ID")
                lu = self._get_lookup(col)
                out[col] = pd.Series([(None if is_missing(v) else lu.setdefault(str(v).strip(), f"{prefix}_{len(lu)+1}")) for v in out[col]], index=out.index)
            elif strat == "scale_numeric":
                factor = float(p.get("factor") if p.get("factor") not in (None, "") else round(0.5 + secrets.randbelow(1500) / 1000, 4))
                self.map["_shared_lookups"].setdefault(f"__factor__{col}", {"value": factor})
                numeric, _ = coerce_numeric_smart(out[col])
                mask_ok = numeric.notna()
                res = out[col].copy().astype(object)
                res[mask_ok] = numeric[mask_ok] * factor
                res[out[col].map(is_missing)] = None
                out[col] = res
            elif strat == "regex_replace":
                pat, pref = p.get("pattern", r"\d+"), p.get("prefix", "N")
                lu, rx = self._get_lookup(col), re.compile(pat)
                out[col] = out[col].map(lambda v: None if is_missing(v) else rx.sub(lambda m: lu.setdefault(m.group(0), f"{pref}_{len(lu)+1}"), str(v)))
            elif strat == "hash_deterministic":
                length, lu = int(p.get("length", 10)), self._get_lookup(col)
                out[col] = pd.Series([(None if is_missing(v) else lu.setdefault(str(v).strip(), hashlib.sha256(str(v).strip().encode()).hexdigest()[:length])) for v in out[col]], index=out.index)
            elif strat == "date_shift":
                days = int(p.get("days") if p.get("days") not in (None, "") else secrets.randbelow(2000) - 1000)
                self.map["_shared_lookups"].setdefault(f"__date_shift__{col}", {"days": days})
                fmt = detect_date_format(out[col])
                parsed = parse_dates_robust(out[col], hint_format=fmt)
                shifted = parsed + pd.Timedelta(days=days)
                out[col] = pd.Series([(None if is_missing(o) else (o if pd.isna(s) else (s.strftime(fmt) if fmt else s))) for o, s in zip(out[col], shifted)], index=out.index)
        for col, cfg in config.items():
            if cfg["strategy"] == "recalc_formula" and cfg.get("params", {}).get("formula", "").strip():
                out[col] = eval_formula(cfg["params"]["formula"], out)
        return out

    def reverse_sheet(self, df: pd.DataFrame, sheet: str) -> pd.DataFrame:
        out = df.copy()
        for col, entry in self.map["sheets"].get(sheet, {}).items():
            if col not in out.columns: continue
            strat = entry["strategy"]
            if strat in ("sequential_id", "hash_deterministic"):
                inv = {v: k for k, v in self._get_lookup(col).items()}
                out[col] = out[col].map(lambda v: None if is_missing(v) else inv.get(v, v))
            elif strat == "regex_replace":
                inv = {v: k for k, v in self._get_lookup(col).items()}
                if inv:
                    rx = re.compile("|".join(re.escape(v) for v in inv))
                    out[col] = out[col].map(lambda v: None if is_missing(v) else rx.sub(lambda m: inv[m.group(0)], str(v)))
            elif strat == "scale_numeric":
                shared = self.map["_shared_lookups"].get(f"__factor__{col}")
                if shared and "value" in shared:
                    factor = float(shared["value"])
                    num, _ = coerce_numeric_smart(out[col])
                    res = out[col].copy().astype(object)
                    mask = num.notna()
                    res[mask] = num[mask] / factor
                    res[out[col].map(is_missing)] = None
                    out[col] = res
            elif strat == "date_shift":
                shared = self.map["_shared_lookups"].get(f"__date_shift__{col}")
                if shared and "days" in shared:
                    days, fmt = int(shared["days"]), self.map["_date_formats"].get(f"{sheet}::{col}")
                    parsed = parse_dates_robust(out[col], hint_format=fmt)
                    back = parsed - pd.Timedelta(days=days)
                    out[col] = pd.Series([(None if is_missing(o) else (o if pd.isna(b) else (b.strftime(fmt) if fmt else b))) for o, b in zip(out[col], back)], index=out.index)
        return out

# ============================================================
#  INTERFACE STREAMLIT
# ============================================================

def init_session():
    for k in ["df_dict", "config", "map_data", "preview_data", "sheet_names", "file_name"]:
        if k not in st.session_state:
            st.session_state[k] = None

init_session()

st.title("🛡️ Pseudonymiseur Expert")
st.markdown("Interface optimisée : modifiez les stratégies dans le panneau de gauche et visualisez instantanément le résultat à droite, sans perte de contexte.")

# 1. Barre supérieure : Chargement
with st.container():
    col_up1, col_up2 = st.columns(2)
    with col_up1:
        uploaded_excel = st.file_uploader("Fichier Excel complet (.xlsx)", type=["xlsx"])
    with col_up2:
        uploaded_map = st.file_uploader("Fichier Map (.json) - Optionnel", type=["json"])

    if uploaded_excel and (st.session_state.file_name != uploaded_excel.name):
        df_dict = pd.read_excel(uploaded_excel, sheet_name=None, dtype=str)
        st.session_state.df_dict = df_dict
        st.session_state.sheet_names = list(df_dict.keys())
        st.session_state.file_name = uploaded_excel.name
        st.session_state.config = {s: {c: {"strategy": "keep", "params": {}} for c in df.columns} for s, df in df_dict.items()}
        st.rerun()

    if uploaded_map:
        try:
            st.session_state.map_data = json.load(uploaded_map)
        except Exception:
            st.error("Fichier Map JSON invalide.")

if st.session_state.df_dict:
    st.divider()
    
    # 2. Layout principal : GAUCHE (Configurations persistantes) / DROITE (Aperçus sticky)
    # Ajustement des proportions pour que le bloc gauche fasse la même largeur visuelle que les options de droite
    col_gauche, col_droite = st.columns([1, 2], gap="large")

    sheet = st.sidebar.selectbox("📂 Sélection d'onglet Excel", st.session_state.sheet_names)
    df_current = st.session_state.df_dict[sheet]

    # == COLONNE GAUCHE : Configuration par ligne (Scrollable) ==
    with col_gauche:
        # --- ACTIONS GLOBALES ALIGNÉES AVEC STRATÉGIES ---
        st.subheader("🚀 Exécution Totale & Export")
        if st.button("🔒 Lancer la Pseudonymisation totale", use_container_width=True, type="primary"):
            with st.spinner("Traitement du fichier entier..."):
                final_engine = PseudoEngine(copy.deepcopy(st.session_state.map_data))
                processed = final_engine.apply_workbook(st.session_state.df_dict, st.session_state.config)
                out_excel, out_map = io.BytesIO(), io.StringIO()
                with pd.ExcelWriter(out_excel, engine="openpyxl") as w:
                    for k, v in processed.items(): v.to_excel(w, sheet_name=k, index=False)
                json.dump(final_engine.map, out_map, indent=2, ensure_ascii=False)
                st.session_state.final_excel = out_excel.getvalue()
                st.session_state.final_map = out_map.getvalue().encode("utf-8")
                st.success("Validé !")

        if "final_excel" in st.session_state:
            dl_col1, dl_col2 = st.columns(2)
            with dl_col1: st.download_button("📥 Excel", data=st.session_state.final_excel, file_name="pseudonymise_output.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            with dl_col2: st.download_button("📥 Map", data=st.session_state.final_map, file_name="map_pseudo.json", mime="application/json", use_container_width=True)

        st.divider()

        st.subheader("⚙️ Stratégies")
        st.caption("Modifiez les actions colonne par colonne.")
        
        # Astuce UI: On utilise st.container(height=...) pour faire une barre de scroll interne !
        # L'utilisateur scroll ce bloc sans perde la vue droite.
        scroll_config = st.container(height=650)
        
        with scroll_config:
            for c_name in df_current.columns:
                cdict = st.session_state.config[sheet][c_name]
                
                # Une carte visuelle par colonne
                with st.expander(f"🔠 {c_name} — Actuel : {STRATEGIES[cdict['strategy']].split('(')[0]}", expanded=False):
                    
                    new_strat = st.selectbox(
                        "Action",
                        options=list(STRATEGIES.keys()),
                        format_func=lambda x: STRATEGIES[x],
                        index=list(STRATEGIES.keys()).index(cdict["strategy"]),
                        key=f"strat_{sheet}_{c_name}"
                    )
                    
                    new_params = {}
                    for p_name, p_label, p_default, p_type in STRATEGY_PARAMS[new_strat]:
                        actual_default = (c_name[:4].upper() if len(c_name) > 0 else "ID") if p_name == "prefix" else p_default
                        val = st.text_input(
                            p_label, 
                            value=str(cdict["params"].get(p_name, actual_default)),
                            key=f"param_{sheet}_{c_name}_{p_name}"
                        )
                        new_params[p_name] = val
                    
                    # Bouton update direct 
                    if st.button("Sauvegarder et rafraîchir", key=f"btn_{sheet}_{c_name}"):
                        st.session_state.config[sheet][c_name] = {"strategy": new_strat, "params": new_params}
                        st.rerun()

    # == COLONNE DROITE : Aperçu (Live) ==
    with col_droite:
        st.subheader("👁️ Aperçu Live")
        
        st.caption("Filtrez et naviguez vos données d'origine de la même manière que sur Excel.")
        with st.expander("🔍 Filtres Excel (Sélectionnez les valeurs exactes à cibler)"):
            filter_cols = st.multiselect("Voulez-vous filtrer par quelles colonnes ?", df_current.columns)
            df_filtered = df_current.copy()
            for col in filter_cols:
                unique_vals = df_filtered[col].dropna().unique().tolist()
                selected_vals = st.multiselect(f"Filtres pour '{col}' :", unique_vals, default=[])
                # Filtration : si rien n'est sélectionné, le tableau devient vide pour cette colonne
                df_filtered = df_filtered[df_filtered[col].isin(selected_vals)]
        
        # Engine execution on the fly
        engine = PseudoEngine(copy.deepcopy(st.session_state.map_data))
        try:
            # Navigation par slider horizontal juste au dessus de la vizu
            max_rows = len(df_filtered)
            if max_rows > 0:
                row_range = st.select_slider(
                    "Navigation : Sélectionnez la plage de lignes à visualiser (Début ↔ Fin)",
                    options=list(range(max_rows + 1)),
                    value=(0, min(100, max_rows))
                )
                df_snippet = df_filtered.iloc[row_range[0]:row_range[1]].copy()
            else:
                df_snippet = df_filtered.copy()
            
            df_preview_full = engine.apply_sheet(df_snippet, sheet, st.session_state.config[sheet])
            
            # --- Flèches de navigation horizontale pour les colonnes ---
            if "col_start" not in st.session_state:
                st.session_state.col_start = 0
            
            max_cols = len(df_preview_full.columns)
            cols_per_page = 5  # Paquet de 5 colonnes
            
            nav_col1, nav_col2, nav_col3 = st.columns([1, 2, 1], vertical_alignment="center")
            with nav_col1:
                if st.button("⬅️ Précédente", disabled=st.session_state.col_start <= 0, use_container_width=True):
                    st.session_state.col_start = max(0, st.session_state.col_start - 1)
                    st.rerun()
            with nav_col2:
                # Ajout d'un slider de navigation fine
                st.session_state.col_start = st.slider(
                    "Défilement colonne par colonne", 
                    0, max(0, max_cols - 1), 
                    st.session_state.col_start,
                    label_visibility="collapsed"
                )
                st.markdown(f"<div style='text-align: center;'>Colonnes <b>{st.session_state.col_start + 1} à {min(st.session_state.col_start + cols_per_page, max_cols)}</b> (sur {max_cols})</div>", unsafe_allow_html=True)
            with nav_col3:
                if st.button("Suivante ➡️", disabled=st.session_state.col_start >= max_cols - 1, use_container_width=True):
                    st.session_state.col_start = st.session_state.col_start + 1
                    st.rerun()
            
            # Application de la pagination des colonnes (Slicing instantané sur tableau pré-calculé)
            df_preview_view = df_preview_full.iloc[:, st.session_state.col_start : st.session_state.col_start + cols_per_page]
            df_snippet_view = df_snippet.iloc[:, st.session_state.col_start : st.session_state.col_start + cols_per_page]
            # -------------------------------------------------------------

            tab_apres, tab_avant = st.tabs(["Résultat (Après)", "Lignes d'origine (Avant)"])
            
            with tab_apres:
                if engine._warnings:
                    for w in engine._warnings: st.warning(w)
                st.dataframe(df_preview_view, use_container_width=True, height=600, hide_index=True)
            with tab_avant:
                st.dataframe(df_snippet_view, use_container_width=True, height=600, hide_index=True)
                
        except Exception as e:
            st.error(f"Erreur empêchant le rendu en direct :\n{e}")
