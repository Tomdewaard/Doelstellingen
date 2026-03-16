"""
Tilstra KPI Watcher — start via START_WATCHER.bat
Leest elke 10 minuten KPI_Invoer.xlsx en schrijft kpi_data.js
De browser herlaadt zichzelf automatisch en toont de nieuwe cijfers.
"""
import sys, os, json, time, re
from pathlib import Path
from datetime import datetime

try:
    import openpyxl
except ImportError:
    print("openpyxl installeren...")
    os.system(f'"{sys.executable}" -m pip install openpyxl')
    import openpyxl

DIR        = Path(__file__).parent
EXCEL      = DIR / "KPI_Invoer.xlsx"
DATA_JS    = DIR / "kpi_data.js"
INTERVAL   = 600  # seconden

KPI_IDS = [
    "sal_stroken","sal_omzet",
    "tr_soll","tr_leads","tr_opdr","tr_start",
    "de_soll","de_kand","de_coll","de_opdr",
    "zzp_leads","zzp_nb",
    "oc_leads","oc_deals","oc_huddle","oc_zzpers",
]

def log(msg):
    print(f"[{datetime.now().strftime('%H:%M:%S')}] {msg}")

def lees_en_schrijf():
    # Excel lezen
    wb = openpyxl.load_workbook(EXCEL, data_only=True)
    ws = wb.active
    waarden = {}
    for i, kid in enumerate(KPI_IDS):
        val = ws.cell(row=5+i, column=3).value
        try:    waarden[kid] = float(val) if val is not None else 0
        except: waarden[kid] = 0
    wb.close()

    # Schrijf kpi_data.js
    tijdstip = datetime.now().strftime("%d-%m-%Y om %H:%M")
    js = f"""// Tilstra KPI Data — automatisch bijgewerkt door watcher.py
// Laatste update: {tijdstip}
window.KPI_DATA = {json.dumps(waarden, indent=2, ensure_ascii=False)};
window.KPI_LAATSTE_UPDATE = "{tijdstip}";
"""
    DATA_JS.write_text(js, encoding="utf-8")
    log(f"✅ kpi_data.js bijgewerkt ({len(waarden)} KPI's)")
    for k,v in waarden.items():
        if v > 0:
            log(f"   {k}: {v}")

def run():
    print("=" * 54)
    print("   Tilstra KPI Watcher")
    print("=" * 54)
    print(f"   Excel:    {EXCEL.name}")
    print(f"   Output:   {DATA_JS.name}")
    print(f"   Interval: elke {INTERVAL//60} minuten")
    print("   Minimaliseer dit venster — sluit om te stoppen.")
    print("=" * 54)
    print()

    if not EXCEL.exists():
        log(f"❌ Niet gevonden: {EXCEL}")
        log("   Zorg dat KPI_Invoer.xlsx in dezelfde map staat.")
        input("Druk Enter om te sluiten...")
        return

    ronde = 0
    while True:
        ronde += 1
        log(f"🔄 Update #{ronde}...")
        try:
            lees_en_schrijf()
        except Exception as e:
            log(f"❌ Fout: {e}")

        # Aftellen
        print()
        for s in range(INTERVAL, 0, -1):
            m = s // 60
            sec = s % 60
            gevuld = int((1 - s/INTERVAL) * 28)
            balk = "█"*gevuld + "░"*(28-gevuld)
            print(f"\r   [{balk}] Volgende update over {m:02d}:{sec:02d} ", end="", flush=True)
            time.sleep(1)
        print()

if __name__ == "__main__":
    run()
