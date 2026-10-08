"""Cambi Banca d'Italia (ex UIC) + costruzione del foglio movimenti in euro."""
from __future__ import annotations

import datetime as dt
import io
import os
import tempfile
from dataclasses import dataclass

import pandas as pd
import requests
from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

BASE = "https://tassidicambio.bancaditalia.it/terzevalute-wf-web/rest/v1.0"
HEADERS = {"Accept": "application/json"}
TIMEOUT = 30
LOOKBACK_DAYS = 10  # copre weekend + festivi consecutivi prima del primo movimento
SOURCE = "Banca d'Italia - tassi di cambio (ex UIC)"
SOURCE_MONTHLY = "Medie mensili Banca d'Italia = provvedimenti Agenzia delle Entrate"
SALDO_TOLERANCE = 0.01


class RatesError(RuntimeError):
    pass


def _get(path: str, **params) -> dict:
    try:
        r = requests.get(f"{BASE}/{path}", params=params, headers=HEADERS, timeout=TIMEOUT)
        r.raise_for_status()
        return r.json()
    except (requests.RequestException, ValueError) as exc:
        raise RatesError(f"Banca d'Italia non raggiungibile o risposta non valida: {exc}") from exc


def _to_units_per_eur(rate: float, convention: str) -> float:
    """Normalizza a 'unità di valuta estera per 1 EUR' (convenzione usata nel foglio)."""
    return 1 / rate if "euro per 1" in convention.lower() else rate


def list_currencies() -> dict[str, str]:
    """{codice ISO: nome} delle valute ancora in vigore."""
    out = {}
    for c in _get("currencies", lang="it")["currencies"]:
        if any(x.get("validityEndDate") is None for x in c["countries"]) and c["isoCode"] != "EUR":
            out[c["isoCode"]] = c["name"]
    return dict(sorted(out.items()))


def daily_rates(iso: str, start: dt.date, end: dt.date) -> pd.DataFrame:
    """Cambi giornalieri: colonne data, cambio (valuta per 1 EUR). Solo giorni con quotazione."""
    data = _get("dailyTimeSeries", startDate=start.isoformat(), endDate=end.isoformat(),
                baseCurrencyIsoCode=iso, currencyIsoCode="EUR", lang="it")
    rows = [(pd.Timestamp(r["referenceDate"]),
             _to_units_per_eur(float(r["avgRate"]), r["exchangeConvention"]))
            for r in data.get("rates", [])]
    if not rows:
        raise RatesError(f"Nessun cambio {iso} tra {start} e {end}.")
    return pd.DataFrame(rows, columns=["data", "cambio"]).sort_values("data").reset_index(drop=True)


def monthly_average(iso: str, year: int, month: int) -> float | None:
    """Cambio medio mensile (stessi valori dei provvedimenti dell'Agenzia delle Entrate)."""
    data = _get("monthlyAverageRates", month=month, year=year,
                baseCurrencyIsoCode=iso, currencyIsoCode="EUR", lang="it")
    rates = data.get("rates") or []
    if not rates:
        return None
    return _to_units_per_eur(float(rates[0]["avgRate"]), rates[0]["exchangeConvention"])


def asof_rate(rates: pd.DataFrame, day: pd.Timestamp) -> tuple[pd.Timestamp, float]:
    """Ultimo cambio con data <= day (weekend/festivi -> giorno precedente)."""
    past = rates[rates["data"] <= day]
    if past.empty:
        raise RatesError(f"Nessun cambio disponibile fino al {day:%d/%m/%Y}.")
    last = past.iloc[-1]
    return last["data"], float(last["cambio"])


@dataclass(frozen=True)
class Mapping:
    data: str
    causale: str | None
    entrate: str | None
    uscite: str | None
    importo: str | None  # alternativa a entrate/uscite: importo con segno
    saldo: str | None


def guess_mapping(columns: list[str]) -> Mapping:
    def find(*keys: str) -> str | None:
        return next((c for c in columns if any(k in str(c).lower() for k in keys)), None)
    split = find("entrat", "avere", "accredit")
    return Mapping(data=find("data", "giorno", "date") or columns[0],
                   causale=find("causale", "descr"),
                   entrate=split, uscite=find("uscit", "dare", "addebit"),
                   importo=None if split else find("importo", "movimento"),
                   saldo=find("saldo"))


def read_table(name: str, content: bytes) -> pd.DataFrame:
    """Legge il primo foglio di un file .xlsx / .ods / .numbers."""
    ext = name.lower().rsplit(".", 1)[-1]
    if ext in ("xlsx", "xlsm"):
        df = pd.read_excel(io.BytesIO(content), engine="openpyxl")
    elif ext == "ods":
        df = pd.read_excel(io.BytesIO(content), engine="odf")
    elif ext == "numbers":
        from numbers_parser import Document
        with tempfile.NamedTemporaryFile(suffix=".numbers", delete=False) as f:
            f.write(content)
        try:
            rows = Document(f.name).sheets[0].tables[0].rows(values_only=True)
        finally:
            os.unlink(f.name)
        df = pd.DataFrame(rows[1:], columns=[str(c) for c in rows[0]])
    else:
        raise ValueError(f"Formato non supportato: .{ext} (usa xlsx, ods o numbers)")
    df = df.dropna(how="all").dropna(axis=1, how="all").reset_index(drop=True)
    df.columns = unique_names(df.columns)
    return df


def unique_names(names) -> list[str]:
    """Nomi colonna univoci: vuoti -> 'colonna N', duplicati -> 'saldo', 'saldo.1', ..."""
    seen: dict[str, int] = {}
    out = []
    for i, n in enumerate(names):
        base = str(n).strip()
        if base in ("", "None", "nan") or base.startswith("Unnamed"):
            base = f"colonna {i + 1}"
        k = seen.get(base, 0)
        seen[base] = k + 1
        out.append(base if k == 0 else f"{base}.{k}")
    return out


def normalize(df: pd.DataFrame, m: Mapping) -> pd.DataFrame:
    """Tabella pulita: data, causale, entrate, uscite (positive), saldo_dichiarato."""
    df = df.loc[:, ~df.columns.duplicated()]  # nomi duplicati: vale la prima colonna
    dates = pd.to_datetime(df[m.data], errors="coerce", dayfirst=True)
    if dates.isna().any():
        raise ValueError(f"Date non valide alle righe: {[i + 1 for i in df.index[dates.isna()]]}")
    num = lambda col: pd.to_numeric(df[col], errors="coerce").fillna(0.0) if col else pd.Series(0.0, index=df.index)
    if m.importo:
        amt = num(m.importo)
        entrate, uscite = amt.clip(lower=0), (-amt).clip(lower=0)
    else:
        entrate, uscite = num(m.entrate), num(m.uscite)
    out = pd.DataFrame({
        "data": dates,
        "causale": df[m.causale].fillna("") if m.causale else "",
        "entrate": entrate, "uscite": uscite,
        "saldo_dichiarato": pd.to_numeric(df[m.saldo], errors="coerce") if m.saldo else float("nan")})
    return out.sort_values("data", kind="stable").reset_index(drop=True)


def saldo_mismatches(mov: pd.DataFrame) -> list[int]:
    """Righe (1-based) in cui il saldo dichiarato non torna con entrate-uscite cumulate."""
    running = (mov["entrate"] - mov["uscite"]).cumsum()
    diff = (mov["saldo_dichiarato"] - running).abs()
    return [i + 1 for i, d in enumerate(diff) if pd.notna(d) and d > SALDO_TOLERANCE]


def months_needed(mov: pd.DataFrame, opening: tuple[int, int]) -> list[tuple[int, int]]:
    """Mese del saldo iniziale + gennaio e dicembre dell'anno dell'ultimo movimento."""
    y = mov["data"].iloc[-1].year
    return sorted({opening, (y, 1), (y, 12)})


def build_workbook(mov: pd.DataFrame, rates: pd.DataFrame, iso: str, opening: tuple[int, int],
                   monthly: dict[tuple[int, int], float | None]) -> bytes:
    """Foglio 'Movimenti' (con formule) + 'Cambi' (giorno per giorno) + 'Medie mensili'."""
    missing = [k for k, v in monthly.items() if v is None]
    if missing:
        raise RatesError(f"Cambio medio mensile non ancora pubblicato per: {missing}")
    wb = Workbook()
    ws, wr, wm = wb.active, wb.create_sheet("Cambi"), wb.create_sheet("Medie mensili")
    ws.title = "Movimenti"
    bold = Font(bold=True)

    wr.append(["data", f"cambio {iso} per 1 EUR", "fonte"])
    for _, r in rates.iterrows():
        wr.append([r["data"].to_pydatetime(), r["cambio"], SOURCE])
    n_rates = wr.max_row

    wm.append(["anno", "mese", f"cambio medio {iso} per 1 EUR", "fonte"])
    keys = sorted(monthly)
    for y, mth in keys:
        wm.append([y, mth, monthly[(y, mth)], SOURCE_MONTHLY])
    opening_row = 2 + keys.index(opening)

    # Layout identico al foglio di riferimento: A..G in valuta, I..N in euro (H vuota).
    ws.append(["data", "causale", "entrate", "uscite", "saldo", "gg", "saldo x gg", None,
               "cambio", "entrate in euro", "uscite in euro", "saldo in euro", "gg", "saldo x gg in euro",
               None, "data cambio usata"])
    n = len(mov)
    lookup = lambda x, col: (f"=LOOKUP(A{x},Cambi!$A$2:$A${n_rates},Cambi!${col}$2:${col}${n_rates})")
    for i, r in mov.iterrows():
        x, first = i + 2, i == 0
        ws.append([
            r["data"].to_pydatetime(), r["causale"], r["entrate"] or None, r["uscite"] or None,
            f"=C{x}-D{x}" if first else f"=E{x-1}+C{x}-D{x}",
            f"=DAYS(A{x+1},A{x})" if i < n - 1 else 0,
            f"=F{x}*E{x}", None,
            f"='Medie mensili'!C{opening_row}" if first else lookup(x, "B"),
            f"=C{x}/I{x}" if r["entrate"] else None,
            f"=D{x}/I{x}" if r["uscite"] else None,
            f"=J{x}-K{x}" if first else f"=L{x-1}+J{x}-K{x}",
            f"=F{x}", f"=M{x}*L{x}", None,
            f"media mensile {opening[1]:02d}/{opening[0]}" if first
            else asof_rate(rates, r["data"])[0].to_pydatetime()])  # data fissa, visibile anche senza ricalcolo
    for row in ws.iter_rows(min_row=2, min_col=1, max_col=1):
        row[0].number_format = "dd/mm/yyyy"
    for row in ws.iter_rows(min_row=2, min_col=16, max_col=16):
        row[0].number_format = "dd/mm/yyyy"
    for row in wr.iter_rows(min_row=2, max_col=1):
        row[0].number_format = "dd/mm/yyyy"
    formats = {"F": "0", "M": "0",  # giorni: numero intero, non una data/durata
               **dict.fromkeys("CDEG", "#,##0.00"),  # valuta: migliaia e centesimi
               **dict.fromkeys("JKLN", "#,##0.000"),  # euro: 3 decimali
               "I": "0.00000"}
    for col, fmt in formats.items():
        for row in ws.iter_rows(min_row=2, min_col=ws[f"{col}1"].column, max_col=ws[f"{col}1"].column):
            row[0].number_format = fmt
    for row in wr.iter_rows(min_row=2, min_col=2, max_col=2):
        row[0].number_format = "0.00000"
    for row in wm.iter_rows(min_row=2, min_col=3, max_col=3):
        row[0].number_format = "0.00000"
    for sheet in (ws, wr, wm):
        for col in range(1, sheet.max_column + 1):
            sheet.cell(row=1, column=col).font = bold
            sheet.column_dimensions[get_column_letter(col)].width = 18

    last = n + 1
    ws.cell(row=last + 2, column=2, value="giacenza media (EUR)").font = bold
    giacenza = ws.cell(row=last + 2, column=14, value=f"=SUM(N2:N{last})/SUM(M2:M{last})")
    giacenza.number_format = "#,##0.000"
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def trim_rates(rates: pd.DataFrame, first_day: pd.Timestamp) -> pd.DataFrame:
    """Tiene i cambi dal 1° gennaio dell'anno di first_day (o dall'ultimo cambio che serve a first_day)."""
    needed = asof_rate(rates, first_day)[0]
    start = min(needed, pd.Timestamp(first_day.year, 1, 1))
    return rates[rates["data"] >= start].reset_index(drop=True)


def convert(mov: pd.DataFrame, iso: str, opening: tuple[int, int]) -> tuple[bytes, pd.DataFrame]:
    """Scarica i cambi necessari e costruisce l'xlsx. Ritorna (xlsx, tabella cambi)."""
    start = mov["data"].iloc[0].date() - dt.timedelta(days=LOOKBACK_DAYS)
    end = mov["data"].iloc[-1].date()
    # la prima riga è il saldo iniziale (cambio medio): i cambi giornalieri servono dalla seconda in poi
    first_day = mov["data"].iloc[1 if len(mov) > 1 else 0]
    rates = trim_rates(daily_rates(iso, start, end), first_day)
    monthly = {k: monthly_average(iso, *k) for k in months_needed(mov, opening)}
    return build_workbook(mov, rates, iso, opening, monthly), rates


def calendar_rates(iso: str, start: dt.date, end: dt.date) -> pd.DataFrame:
    """Un cambio per ogni giorno di calendario: data, cambio, data_quotazione (ultima ufficiale disponibile)."""
    if start > end:
        raise ValueError("La data iniziale è dopo quella finale.")
    quoted = daily_rates(iso, start - dt.timedelta(days=LOOKBACK_DAYS), end)
    days = pd.DataFrame({"data": pd.date_range(start, end)})
    out = pd.merge_asof(days, quoted.rename(columns={"data": "data_quotazione"}),
                        left_on="data", right_on="data_quotazione")
    if out["cambio"].isna().any():
        raise RatesError(f"Nessun cambio {iso} disponibile prima del {start:%d/%m/%Y}.")
    return out[["data", "cambio", "data_quotazione"]]


def calendar_workbook(iso: str, name: str, days: pd.DataFrame) -> bytes:
    """Solo i cambi giorno per giorno di una valuta, in un foglio."""
    wb = Workbook()
    ws = wb.active
    ws.title = f"Cambi {iso}"
    ws.append(["data", f"cambio {iso} per 1 EUR", "data quotazione ufficiale", "fonte"])
    for _, r in days.iterrows():
        ws.append([r["data"].to_pydatetime(), r["cambio"], r["data_quotazione"].to_pydatetime(), SOURCE])
    ws.append([])
    ws.append([f"{name}. Sabato, domenica e festivi non hanno quotazione: vale l'ultima disponibile "
               "(colonna C)."])
    for row in ws.iter_rows(min_row=2, max_row=len(days) + 1):
        row[0].number_format = row[2].number_format = "dd/mm/yyyy"
        row[1].number_format = "0.00000"
    for col in range(1, 5):
        ws.cell(row=1, column=col).font = Font(bold=True)
        ws.column_dimensions[get_column_letter(col)].width = 26
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
