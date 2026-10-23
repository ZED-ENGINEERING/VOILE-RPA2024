# -*- coding: utf-8 -*-
"""
Voile RPA 2024 — squelette d'application Streamlit
==================================================
Interface web au-dessus du moteur de calcul de ``voile_rpa2024.py`` (placer les deux fichiers dans le même dossier).
Aucun calcul n'est dupliqué : l'application ne fait que saisir les données, appeler le moteur et afficher les résultats.

Lancement :   pip install streamlit pandas matplotlib reportlab python-docx openpyxl
              streamlit run streamlit_app.py

Organisation (une fonction ``tab_*`` par rubrique — c'est là qu'il faut enrichir le squelette) :
    Données sismiques · Matériaux · Données générales · Étages / ETABS · Dimensions (ETABS) · Enveloppes M / V ·
    Contraintes (tous étages) · Moment résistant · Élévation · Rives et âme · Résultats · Schéma · Note de calcul

Principe d'état : chaque donnée du moteur (clé de ``DEFAULTS``) est un widget Streamlit portant la même clé dans
``st.session_state`` ; ``build_P()`` reconstitue le dictionnaire ``P`` attendu par le moteur.
"""
import os
import tempfile

import pandas as pd
import streamlit as st
from matplotlib.figure import Figure

import voile_rpa2024 as v

st.set_page_config(page_title="Voile RPA 2024", page_icon="🏗️", layout="wide")

CHECK_KEYS = {"seis_link", "mat_link", "env_auto", "mr_report", "cad_auto"}     # cases à cocher (0 / 1)
STORY_COLS = [c[0] for c in v.STORY_COLS]                                         # name, h, L, e, N, V2, M3, Pmin


# ---------------------------------------------------------------------------
#  État et widgets
# ---------------------------------------------------------------------------
def init_state():
    for k, val in v.DEFAULTS.items():
        if k not in st.session_state:
            st.session_state[k] = bool(val) if k in CHECK_KEYS else val
    st.session_state.setdefault("stories", [])
    st.session_state.setdefault("etabs", None)
    st.session_state.setdefault("dims", None)


def build_P():
    """Dictionnaire de données du moteur, reconstitué depuis ``st.session_state`` (valeurs non affichées = défauts)."""
    P = {}
    for k, default in v.DEFAULTS.items():
        val = st.session_state.get(k, default)
        if k in CHECK_KEYS:
            P[k] = int(bool(val))
        elif k in v.TEXT_KEYS:
            P[k] = str(val)
        else:
            P[k] = float(val) if not isinstance(default, int) else (int(val) if float(val).is_integer() else float(val))
    P["stories"] = v.stories_to_floats(st.session_state["stories"]) if st.session_state["stories"] else []
    return P


def apply_links():
    """Liaison Données sismiques / Matériaux → Données générales (exécutée avant l'affichage des widgets)."""
    try:
        lv = v.link_values(build_P())
    except (ValueError, KeyError) as e:
        st.session_state["_link_err"] = str(e)
        return set()
    st.session_state["_link_err"] = ""
    for k, val in lv.items():
        st.session_state[k] = val
    return set(lv)


def num(label, key, unit="", locked=(), help=None, container=None):
    c = container or st
    default = v.DEFAULTS[key]
    lab = f"{label} [{unit}]" if unit and unit != "-" else label
    if isinstance(default, int) and not isinstance(default, bool):
        return c.number_input(lab, key=key, step=1, disabled=key in locked, help=help)
    return c.number_input(lab, key=key, format="%g", disabled=key in locked, help=help)


def choice(label, key, options, container=None):
    c = container or st
    if st.session_state.get(key) not in options:
        st.session_state[key] = options[0]
    return c.selectbox(label, options, key=key)


def table_from_rows(rows):
    return pd.DataFrame(rows[1:], columns=rows[0]) if rows else pd.DataFrame()


# ---------------------------------------------------------------------------
#  Rubriques
# ---------------------------------------------------------------------------
def tab_seisme(locked):
    st.subheader("Données sismiques — RPA 2024")
    st.checkbox("Lier à Données générales : R → R ; QF → Qf ; T2 → T2 ; T → T0", key="seis_link")
    c1, c2 = st.columns(2)
    with c1:
        choice("Zone sismique", "sz_zone", list(v.RPA_ZONES))
        choice("Groupe d'importance", "sz_groupe", list(v.RPA_GROUPES))
        choice("Classe de site", "sz_site", list(v.RPA_SITES))
        choice("Système de contreventement", "sz_sys", v.SYS_LABELS)
        num("CT", "sz_ct")
        num("hN — hauteur du bâtiment", "sz_hn", "m")
        if st.button("hN ← Σ h des étages") and st.session_state["stories"]:
            st.session_state["sz_hn"] = round(sum(float(s["h"]) for s in st.session_state["stories"]), 3)
            st.rerun()
        choice("Type de structure (amortissement)", "sz_struct", list(v.RPA_AMORT))
        choice("Remplissage", "sz_remp", ["Léger", "Dense"])
    with c2:
        st.markdown("**Facteur de qualité QF = 1 + ΣPq** (critère observé ?)")
        for i, (cat, lab, pq) in enumerate(v.RPA_PQ, start=1):
            choice(f"({cat}) {lab} — Pq = {pq}", f"sz_pq{i}", ["Oui", "Non"])
    try:
        sp = v.seismic_params(build_P())
        st.info(f"A = {sp['A']}  ·  I = {sp['I']}  ·  S = {sp['S']}  ·  T1 = {sp['T1']} s  ·  T2 = {sp['T2']} s  ·  "
                f"T3 = {sp['T3']} s  \nR = {sp['R']} (système {sp['n_sys']}, catégorie {sp['cat']})  ·  "
                f"QF = {sp['QF']:.2f}  ·  ξ = {sp['xi']} %  ·  η = {sp['eta']:.3f}  ·  T = {sp['T']:.3f} s  ·  "
                f"0,2·A·I = {sp['Sa_min']:.3f} g")
    except ValueError as e:
        st.error(str(e))


def tab_materiaux(locked):
    st.subheader("Matériaux")
    st.checkbox("Lier à Données générales : fc28, fe, γb, γs (et Moment résistant, cas de calcul)", key="mat_link")
    c1, c2, c3 = st.columns(3)
    num("fc28 — Fondation", "mt_fc_fond", "MPa", container=c1)
    num("fc28 — Poteaux, voiles, poutres, dalles", "mt_fc", "MPa", container=c1)
    num("Densité du béton", "mt_dens", "T/m³", container=c1)
    num("fe — Barres HA", "mt_fe_ha", "MPa", container=c2)
    num("fe — Treillis soudés", "mt_fe_tr", "MPa", container=c2)
    num("fe — Aciers doux", "mt_fe_doux", "MPa", container=c2)
    choice("Situation de calcul", "mt_sit", list(v.MAT_SITUATIONS), container=c3)
    num("γb manuel", "mt_gb_m", container=c3)
    num("γs manuel", "mt_gs_m", container=c3)
    try:
        m = v.material_params(build_P())
        st.info(f"Ev = {m['Ev']:,.0f} MPa  ·  Ei = {m['Ei']:,.0f} MPa  ·  γb = {m['gb']}  ·  γs = {m['gs']}  ·  "
                f"fbu = {m['fbu']:.2f} MPa  ·  σs = {m['sigs']:.1f} MPa")
    except ValueError as e:
        st.error(str(e))


def first_level(which):
    """Boutons « valeur du 1er niveau » (callbacks) — équivalent de ``take_first`` de la version Tkinter."""
    S = st.session_state["stories"]
    if not S:
        return
    f = S[0]
    if which == "MED":
        st.session_state["MED"] = round(abs(float(f["M3"])), 2)
    elif which == "N":
        st.session_state["N"] = round(abs(float(f["Pmin"])), 2)
    elif which == "V":
        st.session_state["V"] = round(abs(float(f["V2"])), 2)
    elif which == "Nu":
        st.session_state["mr_Nu"] = round(-float(f["N"]), 2)


def tab_general(locked):
    st.subheader("Données générales")
    if locked:
        st.caption("Champs grisés : valeurs imposées par « Données sismiques » / « Matériaux » (décocher la liaison pour "
                   "les saisir librement).")
    for k in ("projet", "element", "niveau", "auteur", "date"):
        st.text_input(k.capitalize(), key=k)
    for k, lab, u in v.GENERAL_FIELDS:
        c1, c2 = st.columns([3, 2])
        num(lab, k, u, locked, container=c1)
        if k in v.FIRST_BTN:
            c2.button(v.FIRST_BTN[k][0], key=f"btn_{k}", on_click=first_level, args=(k,))
        elif k == "MRD":
            def _mru():
                try:
                    mr = v.compute_mr(build_P())
                    if mr["elu"]:
                        st.session_state["MRD"] = round(mr["elu"]["Mru"], 1)
                except ValueError:
                    pass
            c2.button("← Mru (Moment résistant)", key="btn_mru", on_click=_mru)


def _etabs_upload():
    f = st.file_uploader("Export ETABS « Pier Forces » (xlsx / xlsm / csv)", type=["xlsx", "xlsm", "csv", "txt"],
                         key="up_etabs")
    if f is not None and st.session_state.get("_etabs_name") != f.name:
        with tempfile.NamedTemporaryFile(suffix=os.path.splitext(f.name)[1], delete=False) as t:
            t.write(f.getvalue())
        try:
            st.session_state["etabs"] = v.load_etabs(t.name)
            st.session_state["_etabs_name"] = f.name
            st.session_state["etabs_path"] = f.name
        except Exception as e:  # noqa
            st.error(str(e))
        finally:
            os.unlink(t.name)


def tab_etages(locked):
    st.subheader("Étages / ETABS")
    _etabs_upload()
    data = st.session_state["etabs"]
    if data:
        c1, c2 = st.columns([1, 2])
        pier = c1.selectbox("Pier", data["piers"], key="pier_sel")
        if c2.button("Charger les efforts du pier"):
            st.session_state["stories"] = v.stories_from_etabs(data, pier, build_P())
            st.session_state["pier"] = pier
            st.rerun()
    df = pd.DataFrame(st.session_state["stories"], columns=STORY_COLS + ["sec"]) if st.session_state["stories"] \
        else pd.DataFrame(columns=STORY_COLS + ["sec"])
    st.caption("Niveaux du bas vers le haut — N < 0 = compression. Tableau modifiable ; « sec » = section adoptée "
               "facultative (ex. 16HA20).")
    ed = st.data_editor(df, num_rows="dynamic", key="ed_stories")
    st.session_state["stories"] = ed.fillna({"sec": ""}).fillna(0).to_dict("records")
    c1, c2 = st.columns(2)
    if c1.button("Appliquer lw et bw (Données générales) à tous les niveaux"):
        for s in st.session_state["stories"]:
            s["L"], s["e"] = st.session_state["lw"], st.session_state["bw"]
        st.rerun()
    c2.button("Tout appliquer (MED, N, V, Nu) depuis le 1er niveau",
              on_click=lambda: [first_level(k) for k in ("MED", "N", "V", "Nu")])


def tab_dimensions(locked):
    st.subheader("Dimensions des voiles (ETABS « Pier Section Properties »)")
    f = st.file_uploader("Fichier Pier Section Properties", type=["xlsx", "xlsm", "csv", "txt"], key="up_dims")
    if f is not None and st.session_state.get("_dims_name") != f.name:
        with tempfile.NamedTemporaryFile(suffix=os.path.splitext(f.name)[1], delete=False) as t:
            t.write(f.getvalue())
        try:
            st.session_state["dims"] = v.load_pier_sections(t.name)
            st.session_state["_dims_name"] = f.name
        except Exception as e:  # noqa
            st.error(str(e))
        finally:
            os.unlink(t.name)
    d = st.session_state["dims"]
    if not d:
        st.info("Importer un fichier pour calculer L, e et h de chaque pier.")
        return
    T = v.dimension_tables(d)
    st.success(f"{len(d['piers'])} piers, {len(d['stories'])} niveaux — hauteur totale {T['hw_total']:.3f} m — "
               f"Σ surface {T['S_total']:.3f} m²")
    pier = st.selectbox("Pier", d["piers"], key="dim_pier")
    lv = v.pier_levels(d, pier)
    st.dataframe(pd.DataFrame([{"Niveau": r["story"], "Cote base": r["zb"], "Cote sommet": r["zt"], "h": r["h"],
                                "L": r["L"], "e": r["e"], "L×e": round(r["S"], 4)} for r in lv]), hide_index=True)
    c1, c2, c3 = st.columns(3)
    if c1.button("Appliquer h, L, e aux niveaux"):
        n = v.apply_dims_to_stories(st.session_state["stories"], d, pier)
        st.toast(f"{n} niveau(x) mis à jour")
    if c2.button("Créer les niveaux depuis ce pier"):
        st.session_state["stories"] = v.stories_from_dims(d, pier)
        st.rerun()
    if c3.button("lw, bw, he ← niveau le plus bas"):
        st.session_state.update(lw=lv[0]["L"], bw=lv[0]["e"], he=lv[0]["h"])
        st.rerun()
    t1, t2 = st.tabs(["Synthèse par niveau", "Récapitulatif par pier"])
    t1.dataframe(pd.DataFrame(T["levels"]), hide_index=True)
    t2.dataframe(pd.DataFrame(T["recap"]), hide_index=True)
    with tempfile.TemporaryDirectory() as td:
        path = v.export_dimensions_xlsx(d, os.path.join(td, "dimensions.xlsx"))
        st.download_button("Exporter les tableaux (Excel)", open(path, "rb").read(), "Dimensions_voiles_ETABS.xlsx")


def tab_enveloppes(locked, multi):
    st.subheader("Enveloppes des moments et de l'effort tranchant — RPA 2024")
    c1, c2, c3 = st.columns(3)
    c1.checkbox("M base / M sommet automatiques (ETABS)", key="env_auto")
    num("M base", "env_mbase", "kN.m", container=c1)
    num("M sommet", "env_msom", "kN.m", container=c1)
    choice("Hauteur de référence", "env_ref", list(v.REF_LABELS), container=c2)
    mode_keys = list(v.MODE_LABELS)
    st.session_state["m_mode"] = st.session_state["m_mode"] if st.session_state["m_mode"] in mode_keys else "auto"
    c2.selectbox("Moment de calcul", mode_keys, key="m_mode", format_func=v.MODE_LABELS.get)
    choice("Effort tranchant de calcul", "v_mode", list(v.V_LABELS), container=c3)
    if not multi:
        st.info("Aucun étage : importer un export ETABS (rubrique « Étages / ETABS »).")
        return
    env, el, S = multi["env"], multi["elanc"], st.session_state["stories"]
    (st.error if el["elance"] else st.success)(
        f"Art. 7.7.4 — hw/lw = {env['hw']:.2f} / {st.session_state['lw']:.2f} = {el['ratio']:.2f} "
        f"{'>' if el['elance'] else '≤'} {el['lim']} → **{el['cls']}** : {el['regle']}.")
    st.write(f"hcr = {env['hcr']:.3f} m — M base = {env['mbase']:.1f} — M sommet = {env['msom']:.1f} kN.m — "
             f"pente = {env['pente']:.1f} kN.m/m")
    st.dataframe(pd.DataFrame([{"Niveau": "Base" if t["k"] == 0 else S[t["k"] - 1]["name"], "H": t["H"],
                                "M linéaire": t["M_lin"], "M corrigé": t["M_corr"], "Zone": t["zone"]}
                               for t in env["table"]]), hide_index=True)
    st.dataframe(pd.DataFrame([{"Niveau": S[t["k"]]["name"], "H": t["H"], "V analyse": t["Ti"], "1,4·V": t["T14"],
                                "V calcul": t["Vc"], "Zone": t["zone"]} for t in multi["envv"]["table"]]),
                 hide_index=True)


def tab_contraintes(locked, multi):
    st.subheader("Méthode des contraintes — tous étages")
    with st.expander("Paramètres"):
        c = st.columns(5)
        for col, (k, lab) in zip(c, (("kV", "kV"), ("kd", "kd"), ("ktau", "ktau"), ("rho_t", "ρt"), ("rho_g", "ρg"))):
            num(lab, k, container=col)
        choice("Cas de calcul actif", "cas", [st.session_state[f"c{i}_name"] for i in (1, 2, 3)])
    if not multi:
        st.info("Aucun étage.")
        return
    rows = [{"Niveau": r["name"], "N": r["N"], "V2": r["V2"], "M": r["M"], "L×e": f"{r['L']}×{r['e']}",
             "σ1": r["s1"], "σ2": r["s2"], "Type": r["type"], "Lt": r["Lt"], "As calc.": r["As_calc"],
             "As min": r["As_min"], "As requise": r["As"], "Section adoptée": r["sec"],
             "σc/fbu": r["sc_fbu"], "τb/τlim": r["tau_lim"], "Statut": "OK" if r["ok"] else "NG"}
            for r in multi["rows"]]
    st.dataframe(pd.DataFrame(rows), hide_index=True)
    st.caption(f"Cas actif : {multi['cas']['name']} — As requise max = {multi['As_max']:.1f} cm² — "
               f"{multi['n_ng']} niveau(x) NG")


def tab_mr(locked):
    st.subheader("Moment résistant — BAEL 91 / RPA 2024")
    c1, c2, c3 = st.columns(3)
    choice("Situation", "mr_sit", v.MR_SITUATIONS, container=c1)
    for k, lab, u in (("mr_theta", "θ", "-"), ("mr_gb_m", "γb manuel", "-"), ("mr_gs_m", "γs manuel", "-"),
                      ("mr_fc", "fc28", "MPa"), ("mr_fe", "fe", "MPa"), ("mr_Es", "Es", "MPa"), ("mr_n", "n", "-")):
        num(lab, k, u, locked, container=c1)
    for k, lab, u in (("mr_b", "b", "m"), ("mr_h", "h", "m"), ("mr_dp", "d'", "m"), ("mr_c", "c", "m")):
        num(lab, k, u, container=c2)
    num("A' — Armatures supérieures", "mr_Ap", "cm²", container=c2)
    num("A — Armatures inférieures", "mr_A", "cm²", container=c2)
    c2.button("A' ← Avcalc rive gauche ; A ← Avcalc rive droite", key="btn_avc", on_click=lambda: st.session_state.update(
        mr_Ap=st.session_state["Avcalcg"], mr_A=st.session_state["Avcalcd"]))
    choice("Calcul à l'ELU ?", "mr_elu", ["Oui", "Non"], container=c3)
    num("Nu (compression > 0)", "mr_Nu", "kN", container=c3)
    c3.button("Nu ← −N [P] du 1er niveau", key="btn_nu", on_click=first_level, args=("Nu",))
    choice("Calcul à l'ELS ?", "mr_els", ["Oui", "Non"], container=c3)
    num("Ns (compression > 0)", "mr_Ns", "kN", container=c3)
    choice("Fissuration", "mr_fiss", v.MR_FISSURATION, container=c3)
    num("η", "mr_eta", container=c3)
    c3.text_input("Mu (facultatif)", key="mr_Mu")
    c3.text_input("Ms (facultatif)", key="mr_Ms")
    num("νlim", "mr_nulim", container=c3)
    try:
        mr = v.compute_mr(build_P())
    except (ValueError, ZeroDivisionError) as e:
        st.error(str(e))
        return
    for title, rows in mr["sections"][1:]:
        with st.expander(title, expanded=True):
            st.dataframe(pd.DataFrame([{"Résultat": a, "Expression": b, "Valeur": c, "Unité": d}
                                       for a, b, c, d, _ in rows]), hide_index=True)
    fig = Figure(figsize=(5.6, 3.6), dpi=100)
    v.draw_mr(fig, build_P(), mr)
    st.pyplot(fig)


def tab_elevation(locked, res, multi):
    st.subheader("Élévation — condition d'élancement (Art. 7.7.4)")
    if not multi:
        st.info("Aucun étage : importer un export ETABS.")
        return
    fig = Figure(figsize=(8, 8.5), dpi=100)
    v.draw_elevation(fig, build_P(), res, multi)
    st.pyplot(fig)


def tab_rives(locked):
    st.subheader("Rives et âme")
    h = st.columns([3, 2, 2])
    h[1].markdown("**Rive gauche**")
    h[2].markdown("**Rive droite**")
    for k, lab, u in v.SIDE_FIELDS:
        c = st.columns([3, 2, 2])
        c[0].write(f"{lab} [{u}]")
        for col, sd in ((c[1], "g"), (c[2], "d")):
            default = v.DEFAULTS[k + sd]
            if isinstance(default, int):
                col.number_input(k + sd, key=k + sd, step=1, label_visibility="collapsed")
            else:
                col.number_input(k + sd, key=k + sd, format="%g", label_visibility="collapsed")
    st.markdown("**Âme (partie courante)**")
    for k, lab, u in v.WEB_FIELDS:
        num(lab, k, u)
    st.caption("À compléter : table des cadres (type / nombre / longueur) et cumul L → champ « L ».")   # TODO


def tab_resultats(res):
    st.subheader("Résultats — ductilité locale (Art. 7.7.5)")
    for title, heads, rows in res["sections"]:
        with st.expander(title, expanded=False):
            st.dataframe(pd.DataFrame([{"Désignation": r.label, "Expression": r.expr, heads[0]: r.texts()[0],
                                        heads[1]: r.texts()[1], "Unité": r.unit} for r in rows]),
                         hide_index=True)


def tab_schema(res):
    st.subheader("Schéma — plan et élévation de la zone ductile")
    fig = Figure(figsize=(9, 7), dpi=100)
    v.draw_schema(fig, build_P(), res, "portrait")
    st.pyplot(fig)


def tab_note(P, res, multi):
    st.subheader("Note de calcul")
    st.checkbox("Inclure le moment résistant", key="mr_report")
    c1, c2 = st.columns(2)
    for col, fmt, builder, mime in (
            (c1, "pdf", v.build_pdf, "application/pdf"),
            (c2, "docx", v.build_docx,
             "application/vnd.openxmlformats-officedocument.wordprocessingml.document")):
        with tempfile.TemporaryDirectory() as td:
            path = os.path.join(td, f"Note_calcul.{fmt}")
            try:
                builder(P, res, path, multi)
                col.download_button(f"Télécharger la note ({fmt.upper()})", open(path, "rb").read(),
                                    f"Note_calcul_{P['element'].replace(' ', '_')}.{fmt}", mime)
            except Exception as e:  # noqa
                col.error(f"{fmt.upper()} : {e}")


# ---------------------------------------------------------------------------
#  Projet (enregistrer / ouvrir) — mêmes fichiers .json que la version Tkinter
# ---------------------------------------------------------------------------
def sidebar():
    import json
    sb = st.sidebar
    sb.title("Voile RPA 2024")
    up = sb.file_uploader("Ouvrir un projet (.json)", type=["json"], key="up_proj")
    if up is not None and st.session_state.get("_proj_name") != up.name:
        data = json.load(up)
        for k, val in data.items():
            if k == "stories":
                st.session_state["stories"] = val
            elif k in v.DEFAULTS:
                st.session_state[k] = bool(int(float(val))) if k in CHECK_KEYS else (
                    str(val) if k in v.TEXT_KEYS else float(val) if not isinstance(v.DEFAULTS[k], int) else int(float(val)))
        for k in ("seis_link", "mat_link"):        # projet antérieur : ses valeurs sont conservées
            if k not in data:
                st.session_state[k] = False
        st.session_state["_proj_name"] = up.name
        st.rerun()
    data = {k: str(int(st.session_state[k])) if k in CHECK_KEYS else str(st.session_state[k]) for k in v.DEFAULTS}
    data["stories"] = st.session_state["stories"]
    sb.download_button("Enregistrer le projet (.json)", json.dumps(data, indent=2, ensure_ascii=False),
                       "projet_voile.json", "application/json")
    sb.divider()
    help_url = v.HELP_URL
    if help_url:
        sb.link_button("Aide", help_url)
    else:
        sb.button("Aide", on_click=lambda: st.toast("Le site d'aide n'est pas encore en ligne."))
    with sb.expander("Contact"):
        c = v.CONTACT
        st.markdown(f"**{c['bureau']}**  \nIngénieur : {c['ingenieur']}  \n{c['mail']}  \n" + "  \n".join(c["tels"]))
        st.link_button("Écrire un e-mail", "mailto:" + c["mail"])


# ---------------------------------------------------------------------------
#  Page
# ---------------------------------------------------------------------------
def main():
    init_state()
    locked = apply_links()
    sidebar()
    st.title("Voile en béton armé — RPA 2024")
    if st.session_state.get("_link_err"):
        st.warning(f"Hypothèses : {st.session_state['_link_err']}")

    res = multi = None
    err = ""
    try:
        P = build_P()
        res = v.compute(P)
        multi = v.compute_multi(P) if P["stories"] else None
    except (ValueError, ZeroDivisionError, KeyError) as e:
        P, err = None, str(e)
    if err:
        st.error(f"Calcul impossible : {err}")
    elif res["all_ok"] and (multi is None or multi["all_ok"]):
        st.success("Toutes les vérifications sont satisfaites.")
    else:
        ko = [n for n, ok in res["checks"] if not ok]
        st.error(f"{len(ko)} vérification(s) de ductilité non satisfaite(s) : " + " ; ".join(ko)
                 + (f" — {multi['n_ng']} niveau(x) NG (contraintes)" if multi and multi["n_ng"] else ""))

    names = ["Données sismiques", "Matériaux", "Données générales", "Étages / ETABS", "Dimensions (ETABS)",
             "Enveloppes M / V", "Contraintes", "Moment résistant", "Élévation", "Rives et âme", "Résultats", "Schéma",
             "Note de calcul"]
    t = st.tabs(names)
    with t[0]:
        tab_seisme(locked)
    with t[1]:
        tab_materiaux(locked)
    with t[2]:
        tab_general(locked)
    with t[3]:
        tab_etages(locked)
    with t[4]:
        tab_dimensions(locked)
    with t[5]:
        tab_enveloppes(locked, multi)
    with t[6]:
        tab_contraintes(locked, multi)
    with t[7]:
        tab_mr(locked)
    with t[8]:
        tab_elevation(locked, res, multi) if res else st.info("Corriger les données.")
    with t[9]:
        tab_rives(locked)
    with t[10]:
        tab_resultats(res) if res else st.info("Corriger les données.")
    with t[11]:
        tab_schema(res) if res else st.info("Corriger les données.")
    with t[12]:
        tab_note(P, res, multi) if res else st.info("Corriger les données.")


main()
