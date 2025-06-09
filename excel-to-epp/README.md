# Excel-to-EPP Generator

Prosta aplikacja Streamlit, która:
1. Przyjmuje plik Excel z kolumnami `Symbol`, `Nazwa`, `Ilość`, `Cena netto`, `Stawka VAT`
2. Przyjmuje wzorcowy plik EPP
3. Generuje nowy plik EPP z podmienionymi danymi

## Użycie lokalnie

```bash
git clone https://github.com/Kornelint/excel-to-epp.git
cd excel-to-epp
python -m venv venv
# Windows
venv\Scripts\activate
# macOS/Linux
source venv/bin/activate
pip install -r requirements.txt
streamlit run excel_to_epp_app.py
