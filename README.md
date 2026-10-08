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

Stato: prima versione, non ancora validata end-to-end dall'interfaccia.
