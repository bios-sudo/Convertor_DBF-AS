# -*- coding: utf-8 -*-
"""
Aplicatie web BIOSILV Tools - conversii DBF <-> Excel (A/S) si generare fise de teren.

Rulare locala:
    pip install -r requirements.txt
    streamlit run app.py

Deploy gratuit (Streamlit Community Cloud):
    1. Creeaza un repo GitHub cu aceste fisiere (app.py, biosilv_lib.py, fisa_lib.py,
       requirements.txt, FISA_DESCR_default.docx).
    2. Mergi pe https://share.streamlit.io, conecteaza contul GitHub.
    3. "New app" -> alege repo-ul -> main file: app.py -> Deploy.
    4. Primesti un link public (https://<nume>.streamlit.app) pe care il poti
       folosi de pe orice calculator, fara instalare.
"""
import io
import os
import tempfile
import zipfile

import streamlit as st

import biosilv_lib as bl
import fisa_lib

st.set_page_config(page_title="BIOSILV Tools", page_icon="🌲", layout="centered")

st.title("🌲 BIOSILV Tools")
st.caption(
    "Conversie DBF ⇄ Excel (A / S) și generare fișe de teren, "
    "pentru amenajamente silvice."
)

tab_dbf2as, tab_as2dbf, tab_fise = st.tabs(
    ["📤 DBF → A/S (Excel)", "📥 A/S → DBF", "📋 Fișe de teren"]
)

# ---------------------------------------------------------------------------
# TAB 1: DBF -> A.xlsx + S.xlsx
# ---------------------------------------------------------------------------
with tab_dbf2as:
    st.subheader("Convertește un DBF BIOSILV în A.xlsx + S.xlsx")
    st.write(
        "Încarcă fișierul `.DBF` exportat din BIOSILV. Rezultă două fișiere Excel: "
        "**A** (fișa descriptivă, un rând per UA) și **S** (elementele de arboret, "
        "un rând per specie)."
    )

    dbf_file = st.file_uploader("Fișier .DBF", type=["dbf", "DBF"], key="dbf2as_dbf")
    col1, col2 = st.columns(2)
    with col1:
        adm_val = st.number_input("Cod ADM", min_value=0, value=1, step=1, key="dbf2as_adm")
    with col2:
        up_val = st.number_input("Număr UP", min_value=0, value=1, step=1, key="dbf2as_up")

    if st.button("Convertește → A.xlsx + S.xlsx", type="primary", disabled=dbf_file is None):
        with st.spinner("Se convertește..."):
            try:
                with tempfile.TemporaryDirectory() as tmp:
                    dbf_path = os.path.join(tmp, "input.dbf")
                    with open(dbf_path, "wb") as f:
                        f.write(dbf_file.getvalue())
                    a_path = os.path.join(tmp, "A.xlsx")
                    s_path = os.path.join(tmp, "S.xlsx")
                    na, ns = bl.dbf_to_A_S(dbf_path, int(adm_val), int(up_val), a_path, s_path)

                    with open(a_path, "rb") as f:
                        a_bytes = f.read()
                    with open(s_path, "rb") as f:
                        s_bytes = f.read()

                st.success(f"Gata: {na} UA-uri în A.xlsx, {ns} elemente de arboret în S.xlsx.")

                zbuf = io.BytesIO()
                with zipfile.ZipFile(zbuf, "w") as z:
                    z.writestr("A.xlsx", a_bytes)
                    z.writestr("S.xlsx", s_bytes)

                c1, c2, c3 = st.columns(3)
                with c1:
                    st.download_button("⬇️ A.xlsx", a_bytes, file_name="A.xlsx",
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                with c2:
                    st.download_button("⬇️ S.xlsx", s_bytes, file_name="S.xlsx",
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                with c3:
                    st.download_button("⬇️ Ambele (.zip)", zbuf.getvalue(), file_name="A_si_S.zip",
                                        mime="application/zip")
            except Exception as e:
                st.error(f"A apărut o eroare: {e}")

    with st.expander("Notă despre coloanele lipsă din S (vexpr, vexra, vexcu, vexig, vextc)"):
        st.write(
            "Aceste 5 coloane (volume exploatabile pe categorii) nu există în DBF-ul "
            "BIOSILV — sunt calculate separat de softul de amenajament. Vor apărea "
            "goale (0) în S.xlsx."
        )

# ---------------------------------------------------------------------------
# TAB 2: A.xlsx + S.xlsx -> DBF
# ---------------------------------------------------------------------------
with tab_as2dbf:
    st.subheader("Convertește A.xlsx + S.xlsx într-un DBF BIOSILV")
    st.write(
        "Încarcă fișierele **A** și **S**, plus un fișier `.DBF` folosit ca **șablon** "
        "(doar pentru structura de câmpuri — datele lui nu contează). "
        "Ideal, folosește un DBF cât mai apropiat de instalația BIOSILV țintă."
    )

    a_up = st.file_uploader("Fișier A.xlsx", type=["xlsx"], key="as2dbf_a")
    s_up = st.file_uploader("Fișier S.xlsx", type=["xlsx"], key="as2dbf_s")
    tmpl_up = st.file_uploader("DBF șablon (structură de câmpuri)", type=["dbf", "DBF"], key="as2dbf_tmpl")
    out_name = st.text_input("Nume fișier rezultat", value="rezultat.dbf", key="as2dbf_name")

    ready = a_up is not None and s_up is not None and tmpl_up is not None
    if st.button("Convertește → DBF", type="primary", disabled=not ready):
        with st.spinner("Se convertește..."):
            try:
                with tempfile.TemporaryDirectory() as tmp:
                    a_path = os.path.join(tmp, "A.xlsx")
                    s_path = os.path.join(tmp, "S.xlsx")
                    tmpl_path = os.path.join(tmp, "template.dbf")
                    out_path = os.path.join(tmp, "out.dbf")
                    with open(a_path, "wb") as f:
                        f.write(a_up.getvalue())
                    with open(s_path, "wb") as f:
                        f.write(s_up.getvalue())
                    with open(tmpl_path, "wb") as f:
                        f.write(tmpl_up.getvalue())

                    n = bl.a_s_to_dbf(a_path, s_path, tmpl_path, out_path)

                    with open(out_path, "rb") as f:
                        dbf_bytes = f.read()

                st.success(f"Gata: {n} UA-uri scrise în DBF.")
                st.download_button("⬇️ Descarcă DBF", dbf_bytes,
                                    file_name=out_name or "rezultat.dbf",
                                    mime="application/octet-stream")
            except Exception as e:
                st.error(f"A apărut o eroare: {e}")

    with st.expander("De ce am nevoie de un DBF șablon?"):
        st.write(
            "Fișierele DBF folosite de BIOSILV sunt tabele Visual FoxPro legate de o "
            "bază de date (.DBC) și au un bloc special de octeți în antet, specific "
            "acelei baze de date. Pentru ca BIOSILV să recunoască fișierul nou generat, "
            "structura de antet trebuie copiată exact dintr-un fișier DBF real, existent."
        )

# ---------------------------------------------------------------------------
# TAB 3: Fise de teren
# ---------------------------------------------------------------------------
with tab_fise:
    st.subheader("Generează fișele de teren (formular tipizat)")
    st.write(
        "Completează automat **fața** formularului (fișa descriptivă) din A.xlsx + "
        "S.xlsx, și atașează **versoul necompletat** (elemente taxatorice), gata de "
        "dus pe teren. 2 fișe pe pagină, pentru economie de hârtie."
    )

    a_up2 = st.file_uploader("Fișier A.xlsx", type=["xlsx"], key="fise_a")
    s_up2 = st.file_uploader("Fișier S.xlsx", type=["xlsx"], key="fise_s")

    use_default_tmpl = st.checkbox("Folosește formularul standard încărcat în aplicație", value=True)
    tmpl_up2 = None
    if not use_default_tmpl:
        tmpl_up2 = st.file_uploader("Formular FISA_DESCR.docx (propriu)", type=["docx"], key="fise_tmpl")

    col1, col2 = st.columns(2)
    with col1:
        isj_val = st.text_input("Cod ISJ", value="01", key="fise_isj")
    with col2:
        os_val = st.text_input("Cod OS", value="01", key="fise_os")

    ready2 = a_up2 is not None and s_up2 is not None and (use_default_tmpl or tmpl_up2 is not None)
    if st.button("Generează fișele de teren", type="primary", disabled=not ready2):
        with st.spinner("Se generează documentul... (poate dura câteva zeci de secunde pentru multe UA-uri)"):
            try:
                if use_default_tmpl:
                    default_path = os.path.join(os.path.dirname(__file__), "FISA_DESCR_default.docx")
                    with open(default_path, "rb") as f:
                        template_bytes = f.read()
                else:
                    template_bytes = tmpl_up2.getvalue()

                progress = st.progress(0.0)
                docx_bytes, n_ua = fisa_lib.generate_fise(
                    a_up2.getvalue(), s_up2.getvalue(), template_bytes,
                    isj=isj_val, os_code=os_val,
                    progress_cb=lambda p: progress.progress(p),
                )
                progress.empty()

                st.success(f"Gata: {n_ua} UA-uri, {(n_ua + 1) // 2 * 2} pagini (față + verso).")
                st.download_button(
                    "⬇️ Descarcă fișele de teren (.docx)", docx_bytes,
                    file_name="fise_teren.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
            except Exception as e:
                st.error(f"A apărut o eroare: {e}")

    with st.expander("Câmpuri care nu pot fi completate automat"):
        st.write(
            "- **Redenumire UA**, **%SUPR**, **Cantitate lucrări executate**: nu există "
            "în A.xlsx/S.xlsx, rămân goale.\n"
            "- **CAT FUNC FCT** (sub-coloanele 1/2/3): maparea exactă nu a fost încă "
            "100% confirmată pentru cazul cu toate cele 3 sub-câmpuri completate.\n"
            "- **CRS** (creștere) în tabelul de specii: poate arăta uneori doar cifra "
            "zecimală, din cauza unei particularități de aliniere a casetelor."
        )

st.divider()
st.caption(
    "Aceste unelte au fost construite prin compararea directă cu fișiere reale "
    "generate de BIOSILV — verifică întotdeauna rezultatul înainte de utilizare oficială."
)
