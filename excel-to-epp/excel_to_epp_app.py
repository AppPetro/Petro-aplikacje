import streamlit as st
import pandas as pd
import re
from datetime import datetime
from zoneinfo import ZoneInfo

# --- 0) Konfiguracja aplikacji ---
st.set_page_config(page_title="Excel to EPP Generator", layout="wide")
st.title("Excel to EPP Generator 🚀")

# --- 1) Plik referencyjny z opakowaniami i wagami ---
REF_PACKAGING_FILE = "excel_informacyjny.xlsx"
try:
    packaging_df = pd.read_excel(REF_PACKAGING_FILE)
except FileNotFoundError:
    st.error(f"Nie znalazłem pliku `{REF_PACKAGING_FILE}`. Umieść go obok `app.py`.")
    st.stop()

# Uporządkuj nagłówki - usuń nowe linie/spacje
packaging_df.columns = (
    packaging_df.columns
    .str.replace(r"\s+", " ", regex=True)
    .str.strip()
)

# Kolumna z EAN: "Kod EAN" → "Symbol"
if "Kod EAN" in packaging_df.columns:
    packaging_df = packaging_df.rename(columns={"Kod EAN": "Symbol"})
elif "Symbol" not in packaging_df.columns:
    st.error(f"`{REF_PACKAGING_FILE}` musi mieć kolumnę 'Kod EAN' lub 'Symbol'.")
    st.stop()

# Kolumna z wagą → "Waga"
if "Waga, kg" in packaging_df.columns:
    packaging_df = packaging_df.rename(columns={"Waga, kg": "Waga"})
elif "Waga" not in packaging_df.columns:
    st.error(f"`{REF_PACKAGING_FILE}` musi mieć kolumnę 'Waga, kg'.")
    st.stop()

# Kolumna z ilością w opakowaniu
if "Ilość w opakowaniu" not in packaging_df.columns:
    st.error(f"`{REF_PACKAGING_FILE}` musi mieć kolumnę 'Ilość w opakowaniu'.")
    st.stop()

# Oczyść i przekonwertuj dane referencyjne
packaging_df["Symbol"] = (
    packaging_df["Symbol"].astype(str)
    .str.replace(r"\.0+$", "", regex=True)
    .str.strip()
)
packaging_df["Ilość w opakowaniu"] = (
    pd.to_numeric(packaging_df["Ilość w opakowaniu"], errors="coerce")
    .fillna(0)
    .astype(int)
)
packaging_df["Waga"] = (
    pd.to_numeric(packaging_df["Waga"], errors="coerce")
    .fillna(0.0)
    .astype(float)
)

# --- 2) Widżety użytkownika ---
doc_type = st.radio(
    "Typ dokumentu:",
    ("ZK", "MM"),
    index=0,
    help="ZK = Zamówienie od Klienta; MM = Przesunięcie Międzymagazynowe"
)
file_label = st.text_input(
    "Nazwa pliku (bez rozszerzenia):",
    help="bez polskich znaków; spacje → _"
)
use_packages = st.radio(
    "Czy przeliczać na opakowania?", ("Nie", "Tak"), index=0
)
order_file = st.file_uploader(
    "Wgraj Excel z zamówieniem (.xlsx lub .xls)", type=["xlsx", "xls"]
)
if not order_file:
    st.info("Proszę wgrać plik z zamówieniem, aby rozpocząć.")
    st.stop()

# --- 3) Detekcja nagłówków w zamówieniu ---
raw = pd.read_excel(order_file, header=None)

synonyms = {
    "symbol": ["kod ean", "symbol", "ean", "kod produktu"],
    "ilość":  ["ilość", "ilosc", "qty", "ilość sztuk zamówiona"]
}
def clean(cell: str) -> str:
    c = str(cell).lower().strip()
    c = re.sub(r"[^\w\sąćęłńóśźż]", "", c)
    return re.sub(r"\s+", " ", c)

proc = {k: [clean(v) for v in vals] for k, vals in synonyms.items()}
required = {"symbol", "ilość"}
header_row = None
mapping = {}

for idx, row in raw.iterrows():
    cells  = row.astype(str).tolist()
    cleaned = [clean(c) for c in cells]
    m = {}
    for key, alts in proc.items():
        for alt in alts:
            if alt in cleaned:
                m[key] = cells[cleaned.index(alt)].strip()
                break
    if required.issubset(m.keys()):
        header_row, mapping = idx, m
        break

if header_row is None:
    st.error("Nie znaleziono wiersza z kolumnami 'Symbol' i 'Ilość'.")
    st.stop()

st.success(
    f"Znaleziono nagłówki w wierszu {header_row+1}: "
    + ", ".join(f"{k}→{v}" for k, v in mapping.items())
)

order_df = raw.iloc[header_row+1 :].copy()
order_df.columns = raw.iloc[header_row].astype(str).str.strip()
order_df = order_df.rename(columns={
    mapping["symbol"]: "Symbol",
    mapping["ilość"]:  "Ilość"
})

# Oczyść dane zamówienia
order_df["Symbol"] = (
    order_df["Symbol"].astype(str)
    .str.replace(r"\.0+$", "", regex=True)
    .str.strip()
)
order_df["Ilość"] = pd.to_numeric(order_df["Ilość"], errors="coerce")
order_df = order_df.dropna(subset=["Ilość"])
order_df = order_df[order_df["Ilość"] > 0]
if order_df.empty:
    st.error("Brak pozycji z ilością > 0.")
    st.stop()

# --- 4) Merge z danymi referencyjnymi ---
merged = order_df.merge(
    packaging_df[["Symbol", "Ilość w opakowaniu", "Waga"]],
    on="Symbol",
    how="left"
)

# --- 5) Zaokrąglanie do opakowań (opcjonalnie) ---
messages = []
def round_pkg(row):
    pack = int(row.get("Ilość w opakowaniu", 0))
    qty  = int(row["Ilość"])
    if pack > 1:
        rem = qty % pack
        if rem != 0:
            corrected = qty + (pack - rem)
            messages.append(
                f"Poprawiono +{corrected-qty} szt. przy EAN {row['Symbol']}"
            )
            return corrected
    return qty

if use_packages == "Tak":
    merged["Ilość"] = merged.apply(round_pkg, axis=1)
    for msg in messages:
        st.warning(msg)

# --- 6) Wyliczenie wag ---
# waga jednostkowa (kg) i waga całkowita dla każdej pozycji
merged["Waga jednostkowa [kg]"] = merged["Waga"]
merged["Waga całkowita [kg]"] = merged["Waga"] * merged["Ilość"]
# suma wag
total_weight = merged["Waga całkowita [kg]"].sum()

# --- 7) Przygotowanie finalnych kolumn EPP ---
merged["Nazwa"] = ""
merged["Cena netto"] = 0.0
merged["Stawka VAT"] = 8.0
vat5 = {"9120004635976", "9120004635990"}
merged["Stawka VAT"] = merged.apply(
    lambda r: 5.0 if r["Symbol"].rstrip(".") in vat5 else r["Stawka VAT"],
    axis=1
)

data_df = merged[[
    "Symbol",
    "Nazwa",
    "Ilość",
    "Waga jednostkowa [kg]",
    "Waga całkowita [kg]",
    "Cena netto",
    "Stawka VAT"
]]

# --- 8) Generowanie i udostępnienie pliku EPP ---
tpl = "template_ZK.epp" if doc_type == "ZK" else "template_MM.epp"
try:
    lines = open(tpl, encoding="cp1250", errors="ignore").read().splitlines()
except FileNotFoundError:
    st.error(f"Nie znalazłem szablonu `{tpl}`.")
    st.stop()

tags = [ln.strip().upper() for ln in lines]
i1 = tags.index("[NAGLOWEK]")
i2 = tags.index("[ZAWARTOSC]", i1+1)
i3 = tags.index("[NAGLOWEK]", i2+1)
i4 = tags.index("[ZAWARTOSC]", i3+1)
i5 = tags.index("[NAGLOWEK]", i4+1)
i6 = tags.index("[ZAWARTOSC]", i5+1)

now = datetime.now(ZoneInfo("Europe/Warsaw"))
ts_inside = now.strftime("%Y%m%d%H%M%S")
ts_file   = now.strftime("%Y%m%d_%H%M%S")

# nagłówek
lines[i1+1] = (
    lines[i1+1]
    .replace("<typ_dokumentu>", doc_type)
    .replace("<tu_miejscowosc>", '"Wrocław"')
)

part_info       = [ln.replace("<tu_data>", ts_inside) for ln in lines[:i2]]
part_towary_hdr = [ln.replace("<tu_data>", ts_inside) for ln in lines[i3:i4]]
part_cennik_hdr = [ln.replace("<tu_data>", ts_inside) for ln in lines[i5:i6]]

qty_lines = []
for seq, (_, r) in enumerate(data_df.iterrows(), start=1):
    s = r["Symbol"]
    f = [
        str(seq), "1", f'"{s}"', *["0"]*4, *["0.0000"]*2, '"szt."',
        f"{r['Ilość']:.4f}", f"{r['Ilość']:.4f}", *["0.0000"]*3,
        f"{r['Stawka VAT']:.4f}", *["0.0000"]*4, "", ""
    ]
    qty_lines.append(",".join(f))

towary_lines = []
for seq, (_, r) in enumerate(data_df.iterrows(), start=1):
    s = r["Symbol"]
    flds = [""]*42
    flds[0], flds[1] = str(seq), f'"{s}"'
    flds[4] = flds[6] = f'"{r["Nazwa"]}"'
    flds[9] = '"szt."'
    flds[10] = f'"{int(r["Stawka VAT"])}"'
    flds[11] = f"{r['Stawka VAT']:.4f}"
    flds[12], flds[13] = flds[10], flds[11]
    for p in [14,15,23,29,30]:
        flds[p] = "0.0000"
    for p in [17,20,24,27,32,34,35]:
        flds[p] = "0"
    flds[28] = '"szt."'
    towary_lines.append(",".join(flds))

cennik_lines = []
for _, r in data_df.iterrows():
    cennik_lines.append(
        f'"{r["Symbol"]}","Detaliczna",0.0000,0.0000,'
        f'{r["Cena netto"]:.4f},'
        f'{(r["Cena netto"]*(1+r["Stawka VAT"]/100)):.4f},'
        '0.0000'
    )

final = [
    *part_info, "", "[ZAWARTOSC]", *qty_lines,
    "", *part_towary_hdr, "[ZAWARTOSC]", *towary_lines,
    "", *part_cennik_hdr, "[ZAWARTOSC]", *cennik_lines,
]

# usuń podwójne puste
cleaned = []
for ln in final:
    if ln.strip() == "" and cleaned and cleaned[-1].strip() == "":
        continue
    cleaned.append(ln)

epp_content = "\r\n".join(cleaned) + "\r\n"

safe_lbl = re.sub(r"\s+", "_", file_label.strip())
fname = f"{doc_type}_{safe_lbl+'_' if safe_lbl else ''}{ts_file}.epp"

# --- 9) Wyświetlenie całkowitej wagi i przycisk pobrania ---
st.markdown(f"**Łączna waga zamówienia:** {total_weight:.2f} kg")
st.download_button(
    label="Pobierz plik EPP",
    data=epp_content.encode("cp1250"),
    file_name=fname,
    mime="text/plain"
)

# --- 10) Podgląd danych w tabeli ---
st.dataframe(data_df)
