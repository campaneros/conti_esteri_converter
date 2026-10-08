import streamlit as st

import core

MESI = ["gennaio", "febbraio", "marzo", "aprile", "maggio", "giugno", "luglio",
        "agosto", "settembre", "ottobre", "novembre", "dicembre"]

st.set_page_config(page_title="Cambi conto estero", layout="centered")
st.title("Conto estero in euro")
st.write("Carica i movimenti del conto in valuta estera e scarica un foglio Excel "
         "con ogni importo convertito in euro al cambio ufficiale del giorno.")
with st.expander("Come vengono scelti i cambi"):
    st.markdown(
        "- **Cambio del giorno:** quotazione ufficiale della Banca d'Italia (che ha preso il posto "
        "dell'Ufficio italiano cambi) alla data del movimento.\n"
        "- **Sabato, domenica e festivi:** non c'è quotazione, quindi si usa l'ultima disponibile. "
        "Nel foglio trovi la data del cambio effettivamente usato.\n"
        "- **Saldo iniziale:** cambio medio del mese scelto, lo stesso pubblicato ogni mese "
        "dall'Agenzia delle Entrate.\n"
        "- **Risultato:** sempre in euro.")


@st.cache_data(ttl=3600, show_spinner="Carico l'elenco delle valute...")
def currencies():
    return core.list_currencies()


@st.cache_data(ttl=3600, show_spinner="Scarico i cambi dalla Banca d'Italia...")
def convert(mov, iso, opening):
    return core.convert(mov, iso, opening)


st.subheader("1. File con i movimenti")
upload = st.file_uploader("Trascina qui il file o sceglilo dal computer", type=["xlsx", "ods", "numbers"],
                          help="Formati accettati: Excel (.xlsx), OpenDocument (.ods), Numbers (.numbers). "
                               "Servono almeno la data e gli importi in entrata e in uscita.")
if not upload:
    st.stop()

try:
    raw = core.read_table(upload.name, upload.getvalue())
    cur = currencies()
except (ValueError, core.RatesError) as exc:
    st.error(str(exc))
    st.stop()

cols = list(raw.columns)
guess = core.guess_mapping(cols)
none = "(nessuna)"
opt = [none] + cols
pick = lambda label, g, help=None: st.selectbox(label, opt, index=opt.index(g) if g in opt else 0, help=help)

with st.expander("Impostazioni avanzate: dove si trovano i dati nel file"):
    st.caption("Di solito non serve cambiare nulla. Usa queste scelte solo se le colonne del tuo file "
               "hanno nomi diversi da data, causale, entrate, uscite e saldo.")
    c1, c2 = st.columns(2)
    with c1:
        data = st.selectbox("Data del movimento", cols, index=cols.index(guess.data))
        causale = pick("Descrizione (causale)", guess.causale)
        saldo = pick("Saldo dopo il movimento", guess.saldo,
                     help="Solo per controllo: se non torna con gli importi ti avviso.")
    with c2:
        entrate = pick("Importo in entrata", guess.entrate)
        uscite = pick("Importo in uscita", guess.uscite)
        importo = pick("Importo unico con segno", guess.importo,
                       help="Usalo al posto di entrata/uscita se il file ha una sola colonna con "
                            "importi positivi e negativi.")

st.subheader("2. Valuta del conto")
iso = st.selectbox("Valuta", list(cur), format_func=lambda k: f"{k} - {cur[k]}", index=None,
                   placeholder="Scrivi il nome o il codice, ad esempio dollaro o USD",
                   label_visibility="collapsed")
if iso is None:
    st.info("Scegli la valuta in cui è tenuto il conto.")
    st.stop()

none_to_none = lambda v: None if v == none else v
mapping = core.Mapping(data, none_to_none(causale), none_to_none(entrate), none_to_none(uscite),
                       none_to_none(importo), none_to_none(saldo))
try:
    mov = core.normalize(raw, mapping)
except ValueError as exc:
    st.error(f"Non riesco a leggere il file: {exc}")
    st.stop()

st.subheader("3. Controlla i movimenti")
first = mov["data"].iloc[0]
months = [(y, m) for y in (first.year - 1, first.year) for m in range(1, 13)]
opening = st.selectbox("Mese del cambio medio per il saldo iniziale", months,
                       index=months.index((first.year, first.month)),
                       format_func=lambda k: f"{MESI[k[1] - 1]} {k[0]}",
                       help="La prima riga del file è trattata come saldo iniziale e convertita "
                            "con il cambio medio di questo mese.")
bad = core.saldo_mismatches(mov)
if bad:
    st.warning(f"Il saldo indicato nel file non torna con entrate e uscite alle righe {bad}. "
               "Controlla il file prima di procedere.")
else:
    st.success(f"{len(mov)} movimenti letti. I saldi del file tornano.")
st.dataframe(mov.rename(columns={"data": "Data", "causale": "Causale", "entrate": "Entrate",
                                 "uscite": "Uscite", "saldo_dichiarato": "Saldo nel file"}),
             use_container_width=True, hide_index=True)

st.subheader("4. Genera il foglio in euro")
if st.button("Calcola", type="primary"):
    try:
        xlsx, rates = convert(mov, iso, opening)
    except core.RatesError as exc:
        st.error(f"Non sono riuscito a ottenere i cambi: {exc}")
        st.stop()
    st.success(f"Fatto. Cambi {iso}/EUR usati dal {rates['data'].min():%d/%m/%Y} "
               f"al {rates['data'].max():%d/%m/%Y}.")
    st.download_button("Scarica il foglio Excel", xlsx, file_name=f"movimenti_{iso}_in_euro.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.caption("Il file contiene tre fogli: Movimenti (con le formule), Cambi (giorno per giorno) "
               "e Medie mensili.")
