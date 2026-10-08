# conti_esteri_converter

Converte in euro i movimenti di un conto estero (xlsx, ods, numbers) usando i cambi
ufficiali della Banca d'Italia (ex Ufficio italiano cambi).

- Cambio del giorno; sabato, domenica e festivi usano l'ultimo cambio disponibile.
- Saldo iniziale: cambio medio mensile (stessi valori dei provvedimenti dell'Agenzia delle Entrate).
- Output: file Excel con fogli `Movimenti` (formule), `Cambi` e `Medie mensili`.

## Uso

    pip install -r requirements.txt
    streamlit run app.py
    python3 test_core.py

Stato: layout e formule verificati contro il foglio di riferimento (python3 test_compare.py, richiede Numbers su macOS); cambio del giorno, saldo iniziale al cambio medio del mese della prima riga.
