import io
import re
import unicodedata
from collections import Counter, defaultdict
from difflib import SequenceMatcher

import cv2
import fitz
import numpy as np
import pytesseract
import streamlit as st

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter


# ============================================================
# CONFIGURAZIONE PAGINA
# ============================================================

st.set_page_config(
    page_title="Calendario PDF → Excel",
    page_icon="⚽"
)

st.title("⚽ Estrazione calendario PDF")

st.write(
    "Carica il PDF del calendario, scegli la squadra "
    "e scarica automaticamente il file Excel."
)


# ============================================================
# FUNZIONI GENERALI
# ============================================================

def norm(testo):
    """
    Pulisce il testo letto dall'OCR.
    """
    testo = testo or ""

    testo = unicodedata.normalize("NFKD", testo)

    testo = "".join(
        c for c in testo
        if not unicodedata.combining(c)
    )

    testo = testo.upper()

    testo = testo.replace("’", "'")
    testo = testo.replace("`", "'")

    testo = re.sub(
        r"[^A-Z0-9'./ -]+",
        " ",
        testo
    )

    testo = re.sub(
        r"\s+",
        " ",
        testo
    )

    return testo.strip(" -")


# ============================================================
# TRASFORMA PAGINA PDF IN IMMAGINE
# ============================================================

def render_page(documento, numero_pagina, zoom=2):

    pagina = documento[numero_pagina]

    pix = pagina.get_pixmap(
        matrix=fitz.Matrix(zoom, zoom),
        alpha=False
    )

    immagine = np.frombuffer(
        pix.samples,
        dtype=np.uint8
    )

    immagine = immagine.reshape(
        pix.height,
        pix.width,
        pix.n
    )

    if pix.n == 4:

        immagine = cv2.cvtColor(
            immagine,
            cv2.COLOR_RGBA2BGR
        )

    else:

        immagine = cv2.cvtColor(
            immagine,
            cv2.COLOR_RGB2BGR
        )

    return immagine


# ============================================================
# OCR
# ============================================================

def ocr(
    immagine,
    psm=7,
    testo_bianco=False
):

    if immagine.size == 0:
        return ""

    grigio = cv2.cvtColor(
        immagine,
        cv2.COLOR_BGR2GRAY
    )

    # Ingrandisce il testo per migliorare OCR
    grigio = cv2.resize(
        grigio,
        None,
        fx=2.5,
        fy=2.5,
        interpolation=cv2.INTER_CUBIC
    )

    if testo_bianco:

        # Testo bianco su sfondo verde
        _, grigio = cv2.threshold(
            grigio,
            175,
            255,
            cv2.THRESH_BINARY
        )

        grigio = cv2.bitwise_not(
            grigio
        )

    else:

        # Testo scuro su sfondo chiaro
        _, grigio = cv2.threshold(
            grigio,
            0,
            255,
            cv2.THRESH_BINARY
            + cv2.THRESH_OTSU
        )

    configurazione = (
        f"--oem 3 --psm {psm}"
    )

    try:

        testo = pytesseract.image_to_string(
            grigio,
            lang="ita",
            config=configurazione
        )

    except pytesseract.TesseractError:

        testo = pytesseract.image_to_string(
            grigio,
            config=configurazione
        )

    return norm(testo)


# ============================================================
# ESTRAZIONE LOCALITA'
# ============================================================

def estrai_localita(campo):

    campo = norm(campo)

    parti = [
        parte.strip()
        for parte in re.split(
            r"\s*-\s*",
            campo
        )
        if parte.strip()
    ]

    if len(parti) > 1:

        return parti[-1]

    return ""


# ============================================================
# LEGGI PAGINA 2
#
# SOCIETA
# CAMPO / LOCALITA
# INDIRIZZO
# ORARIO
# ============================================================

def leggi_societa(pagina):

    altezza, larghezza = pagina.shape[:2]

    # --------------------------------------------------------
    # Coordinate proporzionali delle colonne
    # --------------------------------------------------------

    # SOCIETA
    x_societa_1 = int(
        larghezza * 0.057
    )

    x_societa_2 = int(
        larghezza * 0.250
    )

    # CAMPO / LOCALITA
    x_campo_1 = int(
        larghezza * 0.290
    )

    x_campo_2 = int(
        larghezza * 0.568
    )

    # INDIRIZZO
    x_indirizzo_1 = int(
        larghezza * 0.568
    )

    x_indirizzo_2 = int(
        larghezza * 0.733
    )

    # ORARIO
    x_orario_1 = int(
        larghezza * 0.733
    )

    x_orario_2 = int(
        larghezza * 0.786
    )

    # --------------------------------------------------------
    # Area contenente le 16 società
    # --------------------------------------------------------

    y_inizio = int(
        altezza * 0.207
    )

    y_fine = int(
        altezza * 0.574
    )

    numero_societa = 16

    altezza_riga = (
        y_fine - y_inizio
    ) / numero_societa

    societa = {}

    # --------------------------------------------------------
    # Legge tutte le righe
    # --------------------------------------------------------

    for i in range(
        numero_societa
    ):

        y1 = int(
            y_inizio
            + i * altezza_riga
        )

        y2 = int(
            y_inizio
            + (i + 1)
            * altezza_riga
        )

        # --------------------------
        # SOCIETA
        # --------------------------

        imm_societa = pagina[
            y1:y2,
            x_societa_1:x_societa_2
        ]

        nome = ocr(
            imm_societa,
            psm=7
        )

        # --------------------------
        # CAMPO / LOCALITA
        # --------------------------

        imm_campo = pagina[
            y1:y2,
            x_campo_1:x_campo_2
        ]

        campo = ocr(
            imm_campo,
            psm=7
        )

        # --------------------------
        # INDIRIZZO
        # --------------------------

        imm_indirizzo = pagina[
            y1:y2,
            x_indirizzo_1:x_indirizzo_2
        ]

        indirizzo = ocr(
            imm_indirizzo,
            psm=7
        )

        # --------------------------
        # ORARIO
        # --------------------------

        imm_orario = pagina[
            y1:y2,
            x_orario_1:x_orario_2
        ]

        testo_orario = ocr(
            imm_orario,
            psm=7
        )

        # --------------------------
        # Ricerca HH:MM
        # --------------------------

        risultato_orario = re.search(
            r"\b([01]?\d|2[0-3])[.:]([0-5]\d)\b",
            testo_orario
        )

        if risultato_orario:

            orario = (
                f"{int(risultato_orario.group(1)):02d}:"
                f"{risultato_orario.group(2)}"
            )

        else:

            orario = ""

        # --------------------------
        # Evita righe vuote
        # --------------------------

        if not nome:
            continue

        if nome in [
            "SOCIETA",
            "SOCIETA'"
        ]:
            continue

        localita = estrai_localita(
            campo
        )

        societa[nome] = {

            "campo":
                campo,

            "localita":
                localita,

            "indirizzo":
                indirizzo,

            "orario":
                orario
        }

    return societa


# ============================================================
# CONFRONTO NOMI SQUADRE
# ============================================================

PAROLE_DA_IGNORARE = {

    "ASD",
    "SSD",
    "SRL",
    "ARL",

    "A.S.D.",
    "S.S.D.",
    "S.R.L.",
    "A.R.L."
}


def nome_confronto(nome):

    parole = []

    for parola in norm(
        nome
    ).split():

        if parola not in PAROLE_DA_IGNORARE:

            parole.append(
                parola
            )

    return " ".join(
        parole
    )


def similarita(
    nome1,
    nome2
):

    nome1 = nome_confronto(
        nome1
    )

    nome2 = nome_confronto(
        nome2
    )

    if not nome1 or not nome2:

        return 0

    # Nomi identici
    if nome1 == nome2:

        return 100

    # Esempio:
    #
    # ARCELLASCO
    #
    # ARCELLASCO CITTA DI ERBA

    if (
        len(nome1) >= 5
        and
        (
            nome1 in nome2
            or
            nome2 in nome1
        )
    ):

        return 96

    parole1 = set(
        nome1.split()
    )

    parole2 = set(
        nome2.split()
    )

    if parole1 and parole2:

        jaccard = (
            100
            * len(
                parole1 & parole2
            )
            / len(
                parole1 | parole2
            )
        )

    else:

        jaccard = 0

    sequenza = (
        100
        * SequenceMatcher(
            None,
            nome1,
            nome2
        ).ratio()
    )

    return max(
        jaccard,
        sequenza
    )


# ============================================================
# TROVA SQUADRA DEL CALENDARIO
# NELLA TABELLA SOCIETA
# ============================================================

def trova_squadra(
    testo,
    societa
):

    migliore = None

    punteggio_migliore = 0

    for nome in societa:

        punteggio = similarita(
            testo,
            nome
        )

        if (
            punteggio
            > punteggio_migliore
        ):

            migliore = nome

            punteggio_migliore = (
                punteggio
            )

    # Soglia minima
    if punteggio_migliore >= 60:

        return migliore

    return None


# ============================================================
# RIQUADRI CALENDARIO PAGINA 1
# ============================================================

def riquadri_calendario(
    pagina
):

    altezza, larghezza = (
        pagina.shape[:2]
    )

    # Coordinate del modello originale
    # 1684 x 1191

    colonne = [

        (30, 340),

        (355, 665),

        (680, 990),

        (1005, 1320),

        (1335, 1650)
    ]

    righe = [

        (310, 470),

        (575, 735),

        (840, 1000)
    ]

    risultati = []

    for y0, y1 in righe:

        for x0, x1 in colonne:

            risultati.append(

                (

                    int(
                        x0
                        / 1684
                        * larghezza
                    ),

                    int(
                        y0
                        / 1191
                        * altezza
                    ),

                    int(
                        x1
                        / 1684
                        * larghezza
                    ),

                    int(
                        y1
                        / 1191
                        * altezza
                    )
                )

            )

    return risultati


# ============================================================
# LEGGI DATE GIORNATA
# ============================================================

def leggi_date(
    pagina,
    box
):

    x0, y0, x1, y1 = box

    altezza = pagina.shape[0]

    # Titolo sopra il riquadro
    titolo = pagina[

        max(
            0,
            y0
            - int(
                altezza * 0.065
            )
        ):
        y0,

        x0:x1
    ]

    testo = ocr(
        titolo,
        psm=6,
        testo_bianco=True
    )

    date_trovate = re.findall(

        r"\b"
        r"(\d{1,2})"
        r"[./-]"
        r"(\d{1,2})"
        r"[./-]"
        r"(20\d{2})"
        r"\b",

        testo
    )

    date = []

    for (
        giorno,
        mese,
        anno
    ) in date_trovate:

        data = (
            f"{int(giorno):02d}/"
            f"{int(mese):02d}/"
            f"{anno}"
        )

        if data not in date:

            date.append(
                data
            )

    if len(date) >= 1:

        andata = date[0]

    else:

        andata = ""

    if len(date) >= 2:

        ritorno = date[1]

    else:

        ritorno = ""

    return (
        andata,
        ritorno
    )


# ============================================================
# LEGGI LE 8 PARTITE DI UNA GIORNATA
# ============================================================

def leggi_partite(
    pagina,
    box,
    societa,
    alias
):

    x0, y0, x1, y1 = box

    ritaglio = pagina[
        y0:y1,
        x0:x1
    ]

    altezza, larghezza = (
        ritaglio.shape[:2]
    )

    altezza_riga = (
        altezza / 8
    )

    partite = []

    for i in range(8):

        y_a = int(
            i * altezza_riga
            + altezza_riga * 0.05
        )

        y_b = int(
            (i + 1) * altezza_riga
            - altezza_riga * 0.05
        )

        # --------------------------------
        # SQUADRA CASA
        # --------------------------------

        immagine_casa = ritaglio[

            y_a:y_b,

            int(
                larghezza * 0.02
            ):
            int(
                larghezza * 0.46
            )
        ]

        testo_casa = ocr(
            immagine_casa,
            psm=7
        )

        # --------------------------------
        # SQUADRA OSPITE
        # --------------------------------

        immagine_ospite = ritaglio[

            y_a:y_b,

            int(
                larghezza * 0.54
            ):
            int(
                larghezza * 0.98
            )
        ]

        testo_ospite = ocr(
            immagine_ospite,
            psm=7
        )

        casa = trova_squadra(
            testo_casa,
            societa
        )

        ospite = trova_squadra(
            testo_ospite,
            societa
        )

        if casa and ospite:

            partite.append(
                (
                    casa,
                    ospite
                )
            )

            alias[casa].append(
                testo_casa
            )

            alias[ospite].append(
                testo_ospite
            )

    return partite


# ============================================================
# ANALIZZA PDF
# ============================================================

def analizza_pdf(
    pdf_bytes
):

    documento = fitz.open(
        stream=pdf_bytes,
        filetype="pdf"
    )

    if len(documento) < 2:

        raise ValueError(

            "Il PDF deve contenere almeno "
            "2 pagine: calendario e "
            "tabella società/impianti."

        )

    # Prima pagina
    pagina_calendario = (
        render_page(
            documento,
            0
        )
    )

    # Seconda pagina
    pagina_societa = (
        render_page(
            documento,
            1
        )
    )

    # Legge automaticamente società
    societa = leggi_societa(
        pagina_societa
    )

    if len(societa) < 8:

        raise ValueError(

            "Non riesco a leggere "
            "correttamente la tabella "
            "delle società."

        )

    alias = defaultdict(
        list
    )

    giornate = []

    boxes = riquadri_calendario(
        pagina_calendario
    )

    # --------------------------------------------------------
    # 15 giornate
    # --------------------------------------------------------

    for numero, box in enumerate(
        boxes,
        start=1
    ):

        data_andata, data_ritorno = (
            leggi_date(
                pagina_calendario,
                box
            )
        )

        partite = leggi_partite(

            pagina_calendario,

            box,

            societa,

            alias
        )

        giornate.append(

            {

                "giornata":
                    numero,

                "andata":
                    data_andata,

                "ritorno":
                    data_ritorno,

                "partite":
                    partite
            }

        )

    # --------------------------------------------------------
    # Nome squadra visualizzato
    # come appare nel calendario
    # --------------------------------------------------------

    nomi_visualizzati = {}

    for nome_societa in societa:

        nomi_letti = [

            norm(x)

            for x in alias[
                nome_societa
            ]

            if norm(x)

        ]

        if nomi_letti:

            nome = Counter(
                nomi_letti
            ).most_common(1)[0][0]

            nomi_visualizzati[
                nome_societa
            ] = nome

        else:

            nomi_visualizzati[
                nome_societa
            ] = nome_societa

    return (
        societa,
        giornate,
        nomi_visualizzati
    )


# ============================================================
# INDIRIZZO COMPLETO
# ============================================================

def indirizzo_completo(
    dati
):

    indirizzo = norm(
        dati.get(
            "indirizzo",
            ""
        )
    )

    localita = norm(
        dati.get(
            "localita",
            ""
        )
    )

    if (
        indirizzo
        and
        localita
    ):

        # Esempio:
        #
        # VIA FRANCO LARATTA - LAZZATE

        return (
            f"{indirizzo} - {localita}"
        )

    if indirizzo:

        return indirizzo

    return localita


# ============================================================
# CREA FILE EXCEL
# ============================================================

def crea_excel(

    squadra_scelta,

    societa,

    giornate,

    nomi_visualizzati

):

    righe = []

    for giornata in giornate:

        for (
            casa,
            ospite
        ) in giornata[
            "partite"
        ]:

            if squadra_scelta not in (

                casa,
                ospite

            ):

                continue

            # =================================================
            # ANDATA
            # =================================================

            dati_casa = (
                societa[
                    casa
                ]
            )

            righe.append(

                [

                    giornata[
                        "andata"
                    ],

                    dati_casa[
                        "orario"
                    ],

                    "CAMPIONATO",

                    nomi_visualizzati[
                        casa
                    ],

                    nomi_visualizzati[
                        ospite
                    ],

                    indirizzo_completo(
                        dati_casa
                    )
                ]

            )

            # =================================================
            # RITORNO
            #
            # casa e ospite invertiti
            # =================================================

            dati_casa_ritorno = (
                societa[
                    ospite
                ]
            )

            righe.append(

                [

                    giornata[
                        "ritorno"
                    ],

                    dati_casa_ritorno[
                        "orario"
                    ],

                    "CAMPIONATO",

                    nomi_visualizzati[
                        ospite
                    ],

                    nomi_visualizzati[
                        casa
                    ],

                    indirizzo_completo(
                        dati_casa_ritorno
                    )
                ]

            )

    # ========================================================
    # CREA EXCEL
    # ========================================================

    wb = Workbook()

    ws = wb.active

    ws.title = (
        "Calendario"
    )

    intestazioni = [

        "Data",

        "Ora",

        "Tipo",

        "Squadra casa",

        "Squadra ospite",

        "Indirizzo"
    ]

    ws.append(
        intestazioni
    )

    for riga in righe:

        ws.append(
            riga
        )

    # ========================================================
    # FORMATTAZIONE
    # ========================================================

    for cella in ws[1]:

        cella.font = Font(
            bold=True
        )

        cella.alignment = Alignment(
            horizontal="center"
        )

    larghezze = [

        14,
        10,
        16,
        30,
        30,
        45
    ]

    for i, larghezza in enumerate(
        larghezze,
        start=1
    ):

        ws.column_dimensions[
            get_column_letter(i)
        ].width = larghezza

    ws.freeze_panes = (
        "A2"
    )

    ws.auto_filter.ref = (
        ws.dimensions
    )

    # ========================================================
    # SALVA IN MEMORIA
    # ========================================================

    buffer = io.BytesIO()

    wb.save(
        buffer
    )

    buffer.seek(
        0
    )

    return (
        buffer.getvalue(),
        righe
    )


# ============================================================
# INTERFACCIA STREAMLIT
# ============================================================

pdf_caricato = st.file_uploader(

    "1. Seleziona il file PDF",

    type=[
        "pdf"
    ]

)


# ============================================================
# DOPO CARICAMENTO PDF
# ============================================================

if pdf_caricato is not None:

    st.info(
        f"PDF selezionato: {pdf_caricato.name}"
    )

    # --------------------------------------------------------
    # Analizza PDF
    # --------------------------------------------------------

    with st.spinner(
        "Sto leggendo il PDF e cercando le squadre..."
    ):

        try:

            (
                societa,
                giornate,
                nomi_visualizzati
            ) = analizza_pdf(

                pdf_caricato.getvalue()

            )

        except Exception as errore:

            st.error(
                f"Errore nella lettura del PDF: {errore}"
            )

            st.stop()

    # --------------------------------------------------------
    # Trova solamente le squadre realmente presenti
    # nel calendario
    # --------------------------------------------------------

    squadre_presenti = set()

    for giornata in giornate:

        for (
            casa,
            ospite
        ) in giornata[
            "partite"
        ]:

            squadre_presenti.add(
                casa
            )

            squadre_presenti.add(
                ospite
            )

    squadre_presenti = sorted(

        squadre_presenti,

        key=lambda x:
            nomi_visualizzati.get(
                x,
                x
            )

    )

    # --------------------------------------------------------
    # Controllo
    # --------------------------------------------------------

    if not squadre_presenti:

        st.error(

            "Non sono riuscito a riconoscere "
            "le squadre nel calendario."

        )

        st.stop()

    st.success(

        f"Squadre trovate: "
        f"{len(squadre_presenti)}"

    )

    # --------------------------------------------------------
    # Crea elenco:
    #
    # ARDOR LAZZATE -> codice interno società
    # --------------------------------------------------------

    opzioni = {}

    for id_societa in squadre_presenti:

        nome = nomi_visualizzati.get(

            id_societa,

            id_societa

        )

        opzioni[
            nome
        ] = id_societa


    # ========================================================
    # SCELTA SQUADRA
    # ========================================================

    squadra_visualizzata = st.selectbox(

        "2. Per quale squadra vuoi l'estrapolazione?",

        list(
            opzioni.keys()
        )

    )

    squadra_id = opzioni[
        squadra_visualizzata
    ]


    # ========================================================
    # GENERA EXCEL
    # ========================================================

    if st.button(

        "3. Genera Excel",

        type="primary",

        use_container_width=True

    ):

        excel_bytes, righe = (
            crea_excel(

                squadra_id,

                societa,

                giornate,

                nomi_visualizzati

            )
        )

        # --------------------------------
        # Nome file
        # --------------------------------

        nome_sicuro = re.sub(

            r"[^A-Z0-9]+",

            "_",

            norm(
                squadra_visualizzata
            )

        ).strip(
            "_"
        )

        nome_file = (

            f"Calendario_"
            f"{nome_sicuro}"
            f".xlsx"

        )

        # --------------------------------
        # Salva nello stato Streamlit
        # --------------------------------

        st.session_state[
            "excel"
        ] = excel_bytes

        st.session_state[
            "nome_file"
        ] = nome_file

        st.session_state[
            "numero_gare"
        ] = len(
            righe
        )


    # ========================================================
    # DOWNLOAD EXCEL
    # ========================================================

    if (
        "excel"
        in st.session_state
    ):

        st.success(

            f"Excel pronto. "
            f"Gare estratte: "
            f"{st.session_state['numero_gare']}"

        )

        st.download_button(

            "⬇️ Scarica Excel",

            data=
                st.session_state[
                    "excel"
                ],

            file_name=
                st.session_state[
                    "nome_file"
                ],

            mime=(
                "application/"
                "vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),

            use_container_width=True
        )


# ============================================================
# PIE' DI PAGINA
# ============================================================

st.divider()

st.caption(
    "Formato indirizzo: VIA / PIAZZA / LARGO ... - PAESE"
)
