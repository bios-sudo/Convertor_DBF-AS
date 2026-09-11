# BIOSILV Tools — aplicație web

Aplicație cu 3 funcții, într-o singură pagină:
1. **DBF → A.xlsx + S.xlsx**
2. **A.xlsx + S.xlsx → DBF** (structura standard BIOSILV e integrată — nu mai trebuie încărcat un șablon)
3. **Fișe de teren** (formularul tipizat, completat automat, 2/pagină, verso necompletat, cu opțiune de a adăuga ani la vârstă)

Fișiere incluse:
- `app.py` — interfața (Streamlit)
- `biosilv_lib.py` — conversiile DBF ⇄ Excel
- `fisa_lib.py` — generarea fișelor de teren
- `FISA_DESCR_default.docx` — formularul standard (poate fi înlocuit din interfață)
- `requirements.txt` — bibliotecile necesare

## Varianta 1 — rulare pe calculatorul tău (5 minute, gratuit, fără cont)

Necesită [Python](https://www.python.org/downloads/) instalat (3.9+).

1. Descarcă toate fișierele de mai sus într-un singur folder.
2. Deschide un terminal (cmd/PowerShell pe Windows, Terminal pe Mac) în acel folder.
3. Rulează:
   ```
   pip install -r requirements.txt
   streamlit run app.py
   ```
4. Se deschide automat în browser, la `http://localhost:8501`. Poți folosi aplicația
   oricând, offline, fără instalare de server.

## Varianta 2 — link public, accesibil de pe orice calculator (gratuit)

Foloseste **Streamlit Community Cloud** (nu necesită cunoștințe tehnice, doar un cont GitHub):

1. Creează un cont gratuit pe [github.com](https://github.com) (dacă nu ai deja).
2. Creează un repository nou (poate fi privat) și încarcă în el toate fișierele de mai sus
   (`app.py`, `biosilv_lib.py`, `fisa_lib.py`, `FISA_DESCR_default.docx`, `requirements.txt`).
3. Mergi pe [share.streamlit.io](https://share.streamlit.io) și conectează-te cu contul GitHub.
4. Apasă **"New app"**, alege repository-ul creat, setează fișierul principal la `app.py`,
   apasă **Deploy**.
5. În 1-2 minute primești un link public de forma `https://numele-tau.streamlit.app`,
   pe care îl poți accesa de pe orice calculator/telefon, fără nimic instalat.

> Notă: pe planul gratuit, aplicația "adoarme" după o perioadă de inactivitate și se
> repornește automat (durează ~30 secunde) la următoarea vizită — normal, nu e o eroare.

## Actualizarea formularului sau a regulilor de conversie

- Pentru a schimba formularul standard, înlocuiește `FISA_DESCR_default.docx` (păstrează
  numele) și redeploy.
- Pentru corecturi la maparea câmpurilor, modifică `fisa_lib.py` (fața fișei) sau
  `biosilv_lib.py` (conversia DBF ⇄ Excel) — sunt fișiere Python simple, comentate în română.

## Ce nu poate face automat aplicația (necesită verificare manuală)

- Coloanele `vexpr, vexra, vexcu, vexig, vextc` din S (volume exploatabile) — nu există
  în DBF-ul BIOSILV, apar goale.
- **VOL, CRS, PEX** din tabelul de specii al fișei de teren rămân goale intenționat —
  sunt valori calculate ulterior, după măsurătorile din teren.
- **CAT FUNC FCT** (sub-coloanele 1/2/3) — maparea nu a fost încă 100% confirmată pentru
  cazul cu toate cele 3 sub-câmpuri completate.
- Câteva câmpuri din fișa de teren (zona ALTITUDINE MAX / unele coduri din tabelul de
  specii, la al 2-lea/3-lea element) pot avea ocazional o mică decalare de aliniere,
  din cauza unei diferențe structurale între formularul-șablon propriu și tabelul
  original generat de BIOSILV. Datele sunt corecte; doar poziția exactă în casetă poate
  varia cu 1 caracter în cazuri rare — semnalează-mi exemple concrete și le pot corecta
  punctual.
