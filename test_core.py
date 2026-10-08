"""Run: python3 test_core.py  (rete necessaria solo per il test live)."""
import datetime as dt
import io
import os

import pandas as pd
from openpyxl import load_workbook

import core


def test_asof_weekend_usa_giorno_precedente():
    rates = pd.DataFrame({"data": pd.to_datetime(["2025-01-02", "2025-01-03", "2025-01-06"]),
                          "cambio": [0.83, 0.84, 0.85]})
    day, rate = core.asof_rate(rates, pd.Timestamp("2025-01-05"))  # domenica
    assert (day, rate) == (pd.Timestamp("2025-01-03"), 0.84)


def test_normalize_importo_con_segno_e_controllo_saldo():
    raw = pd.DataFrame({"data": ["10/01/2025", "05/01/2025"], "importo": [-50, 100], "saldo": [50, 100]})
    mov = core.normalize(raw, core.guess_mapping(list(raw.columns)))
    assert mov["entrate"].tolist() == [100, 0] and mov["uscite"].tolist() == [0, 50]
    assert core.saldo_mismatches(mov) == []
    mov.loc[1, "saldo_dichiarato"] = 40
    assert core.saldo_mismatches(mov) == [2]


def test_build_workbook_struttura():
    mov = pd.DataFrame({"data": pd.to_datetime(["2024-12-31", "2025-01-04"]), "causale": ["saldo", ""],
                        "entrate": [100.0, 0.0], "uscite": [0.0, 10.0], "saldo_dichiarato": [100.0, 90.0]})
    rates = pd.DataFrame({"data": pd.to_datetime(["2025-01-03"]), "cambio": [0.8]})
    wb = load_workbook(io.BytesIO(core.build_workbook(mov, rates, "GBP", (2024, 12), {(2024, 12): 0.83, (2025, 1): 0.84})))
    assert wb.sheetnames == ["Movimenti", "Cambi", "Medie mensili"]
    assert wb["Movimenti"]["I2"].value == "='Medie mensili'!C2"
    assert wb["Cambi"]["B2"].value == 0.8


def test_date_cambio_fisse_e_formati_numerici():
    mov = pd.DataFrame({"data": pd.to_datetime(["2024-12-31", "2025-01-04"]), "causale": ["saldo", ""],
                        "entrate": [100.0, 0.0], "uscite": [0.0, 10.0], "saldo_dichiarato": [100.0, 90.0]})
    rates = pd.DataFrame({"data": pd.to_datetime(["2025-01-03"]), "cambio": [0.8]})
    ws = load_workbook(io.BytesIO(core.build_workbook(mov, rates, "GBP", (2024, 12), {(2024, 12): 0.83}))).active
    assert ws["P3"].value == pd.Timestamp("2025-01-03") and ws["P3"].number_format == "dd/mm/yyyy"  # sabato -> venerdì
    assert ws["P2"].value == "media mensile 12/2024"
    assert ws["E3"].number_format == "#,##0.00" and ws["L3"].number_format == "#,##0.000"


def test_colonne_duplicate_non_rompono_normalize():
    assert core.unique_names(["data", "saldo", "saldo", None]) == ["data", "saldo", "saldo.1", "colonna 4"]
    raw = pd.DataFrame([["01/01/2025", 5, 5, 7]], columns=core.unique_names(["data", "entrate", "saldo", "saldo"]))
    mov = core.normalize(raw, core.guess_mapping(list(raw.columns)))
    assert mov["saldo_dichiarato"].tolist() == [5]


def test_numbers_reale_con_colonne_duplicate():
    import os
    path = os.path.join(os.path.dirname(__file__), "CALCOLO CONTO ESTERO.numbers")
    raw = core.read_table(path, open(path, "rb").read())
    mov = core.normalize(raw, core.guess_mapping(list(raw.columns)))
    assert len(mov) == 44 and core.saldo_mismatches(mov) == []


def test_live_gbp_gennaio_2025():
    rates = core.daily_rates("GBP", pd.Timestamp("2025-01-01").date(), pd.Timestamp("2025-01-14").date())
    assert len(rates) == 9  # niente 1/1, sabati e domeniche
    assert str(rates["data"].dtype) == "datetime64[ns]"  # con pandas 3 la precisione varia: il merge_asof si rompeva
    assert core.monthly_average("GBP", 2025, 2) == 0.83071  # = PDF Agenzia Entrate feb 2025


def test_calendar_rates_un_cambio_per_ogni_giorno():
    days = core.calendar_rates("GBP", dt.date(2025, 1, 1), dt.date(2025, 12, 31))
    assert len(days) == 365 and days["cambio"].notna().all()
    first = days.iloc[0]  # 1 gennaio: festivo, vale il 31/12/2024
    assert first["data_quotazione"] == pd.Timestamp("2024-12-31")
    sat = days[days["data"] == "2025-01-04"].iloc[0]
    assert sat["data_quotazione"] == pd.Timestamp("2025-01-03")
    ws = load_workbook(io.BytesIO(core.calendar_workbook("GBP", "Sterlina", days))).active
    assert ws.max_row == 365 + 3 and ws["B2"].number_format == "0.00000" and ws["A2"].number_format == "dd/mm/yyyy"


def test_foglio_cambi_parte_dal_primo_gennaio_senza_dicembre_precedente():
    mov = pd.DataFrame({"data": pd.to_datetime(["2024-12-31", "2025-01-09", "2025-01-11"]), "causale": ["saldo", "", ""],
                        "entrate": [100.0, 10.0, 5.0], "uscite": [0.0, 0.0, 0.0], "saldo_dichiarato": [100.0, 110.0, 115.0]})
    xlsx, rates = core.convert(mov, "USD", (2024, 12))
    assert rates["data"].min() == pd.Timestamp("2025-01-02") and rates["data"].max() == pd.Timestamp("2025-01-10")
    assert load_workbook(io.BytesIO(xlsx))["Cambi"].max_row == len(rates) + 1


def test_pagina_iniziale_ha_il_download_dei_soli_cambi():
    from streamlit.testing.v1 import AppTest
    at = AppTest.from_file(os.path.join(os.path.dirname(os.path.abspath(__file__)), "app.py")).run(timeout=60)
    assert not at.exception
    sel = next(s for s in at.selectbox if s.key == "only_iso")
    assert sel.value is None  # nessuna valuta preselezionata
    sel.select("GBP").run(timeout=60)
    assert not at.exception and not at.error


if __name__ == "__main__":
    for name, fn in list(globals().items()):
        if name.startswith("test_"):
            fn()
            print("ok", name)
