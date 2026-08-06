# MRP Control Tower — מגדל בקרת חוסרים
# גרסה מתוקנת: סנכרון כמויות QPA ומפתחות מערכת (WIP / System Factor * QPA), 
# סידור עץ BOM דינמי מתוך האקסל, ותיקוני תצוגת UI עקבית.

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, date, timedelta
import io
import requests
import json
from supabase import create_client, Client

# ==========================================================
# HELPERS
# ==========================================================
def safe_num(value, default=0.0):
    n = pd.to_numeric(value, errors='coerce')
    if pd.isna(n):
        return default
    return float(n)

# ==========================================================
# CONFIGURATION & SYSTEM FACTORS
# ==========================================================
GITHUB_URL = "https://raw.githubusercontent.com/orenamram-arch/mrp_checking/main/mrp.xlsx"

ASSEMBLY_SYSTEM_FACTORS_FALLBACK = {
    "1096G860-002": 4,
    "1093U447-001": 4,
    "1093M635-003": 16,
    "1096B650-003": 16,
    "1096G880-003": 4
}
ASSEMBLY_SYSTEM_FACTORS = dict(ASSEMBLY_SYSTEM_FACTORS_FALLBACK)

DEFAULT_SUPPLIER_MAP = {
    "6932T100-001": "הרכבה באופק",
    "02015J1R0PBSTR": "פיניקס טכנולוגיות בע\"מ",
    "0402CS-1N0XJLW": "אלינה - הנדסת אלקטרוניקה (2000) בע\"מ",
    "0402CS-2N7XJSW": "אלינה - הנדסת אלקטרוניקה (2000) בע\"מ",
    "0402X105K160SNT": "ג.פוינט אלקטרוניקס בע\"מ",
    "0603LS-122XGRC": "אלינה - הנדסת אלקטרוניקה (2000) בע\"מ",
    "0603LS-821XJRC": "אלינה - הנדסת אלקטרוניקה (2000) בע\"מ",
    "080-1200": "פיניקס טכנולוגיות בע\"מ",
    "08055C105KAT2A": "פיניקס טכנולוגיות בע\"מ",
    "0805ZC106KAT2A": "פיניקס טכנולוגיות בע\"מ",
    "096-0002-0009": "ריי-קיו בע\"מ",
    "10-646402-192N": "BFE",
    "1000B-5001FX": "מאוזר",
    "1022Y326-001": "אנדרי עיבוד שבבי בע\"מ",
    "1037U443-001": "החרט עיבוד שבבי ממוחשב בע\"מ",
    "1045672-1": "ELECTRO ENTERPRISES NEW YORK LLC",
    "1093L092-001": "BFE",
    "1093L211-001": "BFE",
    "1093L395-002": "הרכבה באופק",
    "1093L396-001": "BFE",
    "1093L944-001": "בן חמו אהרון - עיבוד שבבי בע\"מ",
    "1093M635-003": "הרכבה באופק",
    "1093N696-001": "BFE",
    "1093N705-001": "אנדרי עיבוד שבבי בע\"מ",
    "1093R379-001": "כימוגרף-צריבה פוטוכימית בע\"מ",
    "1093U332-001": "פרונת תעשיות בע\"מ",
    "1093U335-001": "פרונת תעשיות בע\"מ",
    "1093U336-001": "פרונת תעשיות בע\"מ",
    "1093U337-001": "פרונת תעשיות בע\"מ",
    "1093U338-001": "גומיאן מוצרי גומי בע\"מ",
    "1093U339-001": "גומיאן מוצרי גומי בע\"מ",
    "1093U347-001": "גומיאן מוצרי גומי בע\"מ",
    "1093U348-001": "גומיאן מוצרי גומי בע\"מ",
    "1093U442-001": "פרונת תעשיות בע\"מ",
    "1093U447-001": "WAVE",
    "1093W210-001": "הרכבה באופק",
    "1093W211-001": "WAVE",
    "1093W212-001": "WAVE",
    "1093W213-001": "WAVE",
    "1094B009-001": "אנדרי עיבוד שבבי בע\"מ",
    "1094G851-001": "שמש גומיאן",
    "1094U399-001": "BFE",
    "1094U750-001": "BFE",
    "1096A410-001": "BFE",
    "1096B634-001": "פרונת תעשיות בע\"מ",
    "1096B650-003": "הרכבה באופק",
    "1096B652-003": "BFE",
    "1096C850-003": "הרכבה באופק",
    "1096C852-002": "BFE",
    "1096D240-003": "הרכבה באופק",
    "1096D242-002": "BFE",
    "1096D244-001": "צמח משה טכנולוגיות 2000 בע\"מ",
    "1096F000-003": "הרכבה באופק",
    "1096F009-003": "אקרמן הדפסות ופוטו אנודייז בע\"מ",
    "1096F010-002": "בן חמו אהרון - עיבוד שבבי בע\"מ",
    "1096F011-002": "בן חמו אהרון - עיבוד שבבי בע\"מ",
    "1096F529-001": "פרונת תעשיות בע\"מ",
    "1096F588-001": "החרט עיבוד שבבי ממוחשב בע\"מ",
    "1096F598-002": "בן חמו אהרון - עיבוד שבבי בע\"מ",
    "1096F602-001": "WAVE",
    "1096G860-002": "הרכבה באופק",
    "1096G862-002": "BFE",
    "1096G880-003": "הרכבה באופק",
    "1096G882-002": "BFE",
    "1096J793-002": "BFE",
    "1096J794-002": "BFE",
    "1096J797-001": "אנדרי עיבוד שבבי בע\"מ",
    "1096J798-001": "אנדרי עיבוד שבבי בע\"מ",
    "1096J800-001": "הרכבה באופק",
    "1096J801-001": "ניק",
    "1096J802-001": "BFE",
    "1096J805-001": "אנדרי עיבוד שבבי בע\"מ",
    "1096J810-001": "הרכבה באופק",
    "1096J811-001": "אנדרי עיבוד שבבי בע\"מ",
    "1096J812-001": "BFE",
    "1096J814-001": "גומיאן מוצרי גומי בע\"מ",
    "1096M790-001": "BFE",
    "1096U730-001": "BFE",
    "1097A328-002": "BFE",
    "1097A531-001": "BFE",
    "1097L001-002": "שירטק בע\"מ",
    "1110-2885-6200": "Winchester Interconnect RF Corporation",
    "145-AA1J/118": "Huang Liang Technologies Co., Ltd",
    "145-AA1R/111": "Huang Liang Technologies Co., Ltd",
    "1857657-2": "ELECTRO ENTERPRISES NEW YORK LLC",
    "18K118-K17L5": "מ.ט.י הנדסה בע\"מ",
    "18S10T-40ML5": "מ.ט.י הנדסה בע\"מ",
    "18S141-40ML5": "מ.ט.י הנדסה בע\"מ",
    "1982257-5": "ARROW",
    "1982260-5": "ריי-קיו בע\"מ",
    "19K101-K00L5": "מ.ט.י הנדסה בע\"מ",
    "19K119-K25D3": "מ.ט.י הנדסה בע\"מ",
    "19S14J-400L5": "מ.ט.י הנדסה בע\"מ",
    "19S14T-40ML5": "מ.ט.י הנדסה בע\"מ",
    "19S16D-400L5": "מ.ט.י הנדסה בע\"מ",
    "19S16G-400L5": "מ.ט.י הנדסה בע\"מ",
    "19S241-S01L5": "מ.ט.י הנדסה בע\"מ",
    "20-41111": "ELECTRO ENTERPRISES NEW YORK LLC",
    "2015D989-001": "אנדרי עיבוד שבבי בע\"מ",
    "2042088-1": "ריי-קיו בע\"מ",
    "2201E366-001": "בן חמו אהרון - עיבוד שבבי בע\"מ",
    "2201E369-001": "אנדרי עיבוד שבבי בע\"מ",
    "2201E370-001": "הרכבה באופק",
    "2201E374-001": "גומיאן מוצרי גומי בע\"מ",
    "2201E375-001": "מיקו כרסום בע\"מ",
    "2201E377-001": "אנדרי עיבוד שבבי בע\"מ",
    "2201E381-001": "אור- דיוק תעשיות בע\"מ",
    "2201E389-001": "פרונת תעשיות בע\"מ",
    "2201E402-003": "BFE",
    "2201E410-005": "הרכבה באופק",
    "2201E412-002": "BFE",
    "2201E414-001": "BFE",
    "2201E427-001": "BFE",
    "2201E440-001": "הרכבה באופק",
    "2201E701-002": "הרכבה באופק",
    "2201E702-002": "BFE",
    "242 50CC": "רוטל דבקים וכימיקלים בע\"מ",
    "262-50CC": "רוטל דבקים וכימיקלים בע\"מ",
    "2N7002LT1G": "ARROW",
    "3-5353652-6": "ריי-קיו בע\"מ",
    "3-907EM163-80": "Hardware Specialty Co., Inc.",
    "30-03-2071-1285": "ELECTRO ENTERPRISES NEW YORK LLC",
    "30-03-2072-1285": "אלימק הנדסה אלקטרו-מכנית (1988) בע\"מ",
    "3769000G001-001": "Delta Electronics Manufacturing Corporation",
    "3SF162-0008-AX": "Huang Liang Technologies Co., Ltd",
    "4HP5ON-SS": "ליא הנדסה",
    "5206F264-001": "אלינה - הנדסת אלקטרוניקה (2000) בע\"מ",
    "5AGXBA5D4F27I5G": "איסטרוניקס בע\"מ",
    "5AGXFB1H4F35I3G": "איסטרוניקס בע\"מ",
    "6930N012-001": "אקרמן הדפסות ופוטו אנודייז בע\"מ",
    "6930N013-001": "אקרמן הדפסות ופוטו אנודייז בע\"מ",
    "6930N014-001": "פוטו מטאל",
    "6930N015-001": "אקרמן הדפסות ופוטו אנודייז בע\"מ",
    "6930N016-001": "פוטו מטאל",
    "6930N127-001": "הרכבה באופק",
    "6930N128-001": "די גי קי",
    "6930N141-001": "הרכבה באופק",
    "6930N142-001": "BFE",
    "6930N150-002": "ויטק ניהול פרוייקטים",
    "6930N195-001": "סינרג'י אר.אם בע\"מ",
    "6930N202-001": "נח-הראל בע\"מ",
    "6930N203-001": "אנדרי עיבוד שבבי בע\"מ",
    "6930N206-001": "ליפין",
    "6930N207-001": "גומיאן מוצרי גומי בע\"מ",
    "6930N216-001": "BFE",
    "6930N217-001": "גומיאן מוצרי גומי בע\"מ",
    "6930N218-001": "גומיאן מוצרי גומי בע\"מ",
    "6930N220-002": "הרכבה באופק",
    "6930N230-002": "הרכבה באופק",
    "6930N232-001": "מיקו כרסום בע\"מ",
    "6930N233-001": "מיקו כרסום בע\"מ",
    "6930N234-001": "רותם עיבוד שבבי",
    "6930N235-001": "WAVE",
    "6930N237-001": "פרונת תעשיות בע\"מ",
    "6930N239-001": "גומיאן מוצרי גומי בע\"מ",
    "6930N240-001": "הרכבה באופק",
    "6930N242-001": "פרונת",
    "6930N243-001": "WAVE",
    "6930N307-001": "אור- דיוק תעשיות בע\"מ",
    "6930N308-001": "אור- דיוק תעשיות בע\"מ",
    "6930N328-001": "עגם תשלובת מפעלים בע\"מ",
    "6930N329-001": "עגם תשלובת מפעלים בע\"מ",
    "6930N348-001": "גומיאן מוצרי גומי בע\"מ",
    "6930N357-001": "BFE",
    "6930N375-001": "WAVE",
    "6932T111-001": "אקרמן הדפסות ופוטו אנודייז בע\"מ",
    "6932T120-001": "הרכבה באופק",
    "6932T122-001": "אקרמן הדפסות ופוטו אנודייז בע\"מ",
    "6932T130-001": "אנרקון טכנולוגיות",
    "6932T200-001": "הרכבה באופק",
    "6932T201-001": "בן חמו אהרון - עיבוד שבבי בע\"מ",
    "6932T205-001": "אקרמן הדפסות ופוטו אנודייז בע\"מ",
    "6932T300-001": "הרכבה באופק",
    "6932T311-001": "אקרמן הדפסות ופוטו אנודייז בע\"מ",
    "6932T335-001": "ק.ש.ר תעשיות פלסטיק",
    "6932T346-001": "ליפין",
    "6932T356-001": "WAVE",
    "6932T450-001": "BFE",
    "74AHC1G125DCKTG4": "PCG Trading- Converge",
    "74LCX244MTCX": "ARROW",
    "8P34S1204NLGI8": "יומינטק",
    "8P34S1208NBGI": "DIGI KEY",
    "A-A-52080-B-3 BLACK": "ELECTRO ENTERPRISES NEW YORK LLC",
    "A-A-52080-E-3 BLACK": "ELECTRO ENTERPRISES NEW YORK LLC",
    "A82.A15GHC.00225": "אס.טי.גי. אינטרנשיונל בע\"מ",
    "ABMM-AT-D": "אלכסנדר שניידר",
    "AD5144BCPZ100-RL7": "אלבטק תעשייה ולוגיסטיקה בע\"מ",
    "AD780BRZ-REEL7": "אלבטק תעשייה ולוגיסטיקה בע\"מ",
    "AD7928BRUZ-REEL7": "אלבטק תעשייה ולוגיסטיקה בע\"מ",
    "AD9253BCPZ-80": "אלבטק תעשייה ולוגיסטיקה בע\"מ",
    "AD9508BCPZ": "אלבטק תעשייה ולוגיסטיקה בע\"מ",
    "ADA4932-2YCPZ-R2": "BFE",
    "ADCLK907BCPZ-R2": "אלבטק תעשייה ולוגיסטיקה בע\"מ",
    "ADCLK925BCPZ-R2": "אלבטק תעשייה ולוגיסטיקה בע\"מ",
    "ADG1634BCPZ-REEL7": "אלבטק תעשייה ולוגיסטיקה בע\"מ",
    "ADG4612BRUZ": "אלבטק תעשייה ולוגיסטיקה בע\"מ",
    "ADRF6780ACPZN": "אלבטק תעשייה ולוגיסטיקה בע\"מ",
    "ADXL375BCCZ": "אלבטק תעשייה ולוגיסטיקה בע\"מ",
    "AF0201JR-0733RL": "קבוצת סידב בע\"מ",
    "AKF-1938": "BFE",
    "AM26LV31EIPWR": "אלבטק תעשייה ולוגיסטיקה בע\"מ",
    "AM26LV31ESDREP": "אלבטק תעשייה ולוגיסטיקה בע\"מ",
    "AM26LV32EIPWR": "אלבטק תעשייה ולוגיסטיקה בע\"מ",
    "AS0B226-S68Q-7H": "פיניקס טכנולוגיות בע\"מ",
    "ASMT-RF45-AN002": "DIGI KEY",
    "ASMT-RR45-AQ902": "DIGI KEY",
    "AT24CM01-XHM-B": "ARROW",
    "AT24CM02-SSHD-B": "אס.טי.גי. אינטרנשיונל בע\"מ",
    "AT24MAC402-SSHM-T": "ARROW",
    "B0540WS-7": "ARROW",
    "B096QC2S-T": "אלינה - הנדסת אלקטרוניקה (2000) בע\"מ | מייקטלוג-70 יח' ל FAI | 100 יח' ל FAI DIGI KEY",
    "BAT54WS": "ARROW",
    "BAT54WS-E3-08": "אקומל ישראל בע\"מ",
    "BD60120N50100AHF": "מאוזר",
    "BLM15AX121SN1D": "ARROW",
    "BLM15EG121SN1D": "אס.טי.גי. אינטרנשיונל בע\"מ",
    "BLM15EG221SN1D": "אס.טי.גי. אינטרנשיונל בע\"מ",
    "BLM15HD182SN1D": "אס.טי.גי. אינטרנשיונל בע\"מ",
    "BLM15HD601SN1D": "אס.טי.גי. אינטרנשיונל בע\"מ",
    "BLM18EG221SN1D": "אס.טי.גי. אינטרנשיונל בע\"מ",
    "BLM18SG260TN1D": "אס.טי.גי. אינטרנשיונל בע\"מ",
    "BLM21PG221SN1D": "אס.טי.גי. אינטרנשיונל בע\"מ",
    "BLM21PG600SN1D": "אס.טי.גי. אינטרנשיונל בע\"מ",
    "BLM31PG330SN1L": "ARROW",
    "BLM31PG601SN1L": "אס.טי.גי. אינטרנשיונל בע\"מ",
    "BP0805A7250ASTR": "שירטק בע\"מ",
    "C0024090": "ניקומטיק",
    "C0201C101K3GACTU": "ARROW",
    "C0402C100C5GAC7867": "מאוזר",
    "C0402C101J3GACTU": "ARROW",
    "C0402C101K5GAC7867": "ARROW",
    "C0402C102K5RAC7867": "ARROW",
    "C0402C103K3RAC7867": "ARROW",
    "C0402C103K3RACTU": "ARROW",
    "C0402C103K5RACTU": "ARROW",
    "C0402C103K5RALTU": "ARROW",
    "C0402C151K5RACTU": "ARROW",
    "C0402C220K5GAC7867": "ARROW",
    "C0402C222K5RACTU": "ARROW",
    "C0402C392K5RACTU": "ARROW",
    "C0402C508C5GACTU": "ARROW",
    "C0603C103K2RAL7867": "מאוזר",
    "C0603C103K5RAC7867": "ARROW",
    "C0603C104K4RACTU": "ARROW",
    "C0603C104K5RAC7867": "ARROW",
    "C0603C105K3RACTU": "ARROW",
    "C0603C106M9PAC7867": "ARROW",
    "C0603C472G5GACTU": "ARROW",
    "C0603C473K3RALTU": "ARROW",
    "C0603C474K4RAL7867": "מאוזר | ARROW",
    "C0805C106K8RAC7800": "ARROW",
    "C0805C225K4RAC7800": "ARROW",
    "C0805C475K4RAC7025": "ARROW",
    "C0805C476M9PAC7800": "ARROW",
    "C10-218": "פלדמן פלד בן ציון - סוכנויות",
    "C1005X7S1A105K050BC": "ARROW",
    "C1206C106K3RACTU": "מייקטלוג",
    "C1210C106K5RACTU": "ARROW",
    "C1210C226M3RACTU": "ARROW",
    "C1210C476M8RACTU": "ARROW",
    "C2012X7R1E475K125AB": "ARROW",
    "CA70P3123HLT": "דונטק אלקטרוניקה בע\"מ",
    "CC1210KKX7R0BB334": "קבוצת סידב בע\"מ",
    "CHT7010088": "דונטק אלקטרוניקה בע\"מ",
    "CL03B102KO3NNNC": "קבוצת סידב בע\"מ | מאוזר",
    "CL03B103KP3NNNC": "קבוצת סידב בע\"מ",
    "CL05B104KP5NNNC": "קבוצת סידב בע\"מ",
    "CL05C010BB5NNNC": "קבוצת סידב בע\"מ",
    "CL10B104KA8NNNC": "קבוצת סידב בע\"מ",
    "CL10B105KO8NNNC": "קבוצת סידב בע\"מ",
    "CL10C470GB8NNNC": "מאוזר",
    "CL10C470JB8NNNC": "קבוצת סידב בע\"מ",
    "CL31A107MQHNNNE": "קבוצת סידב בע\"מ",
    "CL32B106KBJNNWE": "קבוצת סידב בע\"מ",
    "CLH-105-L-D-DV-TR": "SAMTEC Inc.",
    "CRCW02010000Z0ED": "ARROW",
    "CRCW0201100KFKED": "אקומל ישראל בע\"מ",
    "CRCW020110K0FKED": "אקומל ישראל בע\"מ",
    "CRCW02011K00FKED": "אקומל ישראל בע\"מ",
    "CRCW0201200RFKED": "אקומל ישראל בע\"מ",
    "CRCW0201274RFKED": "ARROW",
    "CRCW0201300RFKED": "אקומל ישראל בע\"מ",
    "CRCW020149R9FKED": "אקומל ישראל בע\"מ",
    "CRCW0201887RFKED": "אקומל ישראל בע\"מ",
    "CRCW04020000Z0ED": "ARROW",
    "CRCW04020000ZSED": "ARROW",
    "CRCW0402100KFKED": "ARROW",
    "CRCW0402100RFKED": "ARROW",
    "CRCW040210K0FKED": "ARROW",
    "CRCW040211K0FKED": "ARROW",
    "CRCW040211R5FKED": "מאוזר",
    "CRCW0402121KFKED": "ARROW",
    "CRCW0402121RFKED": "ARROW",
    "CRCW040212K1FKED": "ARROW",
    "CRCW040212R1FKED": "מאוזר",
    "CRCW0402133RFKED": "מאוזר",
    "CRCW040213K3FKED": "ARROW",
    "CRCW0402140KFKED": "ARROW",
    "CRCW040214K7FKED": "ARROW",
    "CRCW0402150RFKED": "ARROW",
    "CRCW040215K0FKED": "ARROW",
    "CRCW0402178RFKED": "ARROW",
    "CRCW040218K2FKED": "ARROW",
    "CRCW040219K1FKED": "ARROW",
    "CRCW04021K00FKED": "ARROW",
    "CRCW04021K21FKED": "ARROW",
    "CRCW04021K47FKED": "ARROW",
    "CRCW04021K50FKED": "ARROW",
    "CRCW04021K82FKTD": "אקומל ישראל בע\"מ",
    "CRCW04021K91FKTD": "אקומל ישראל בע\"מ",
    "CRCW0402200RFKED": "ARROW",
    "CRCW040220K0FKED": "ARROW",
    "CRCW040220K5FKED": "ARROW",
    "CRCW0402221RFKED": "מאוזר",
    "CRCW040222R1FKED": "מאוזר",
    "CRCW0402249RFKED": "ARROW",
    "CRCW040224R3FKED": "ARROW",
    "CRCW040225K5FKED": "ARROW",
    "CRCW0402261RFKED": "מאוזר",
    "CRCW0402274RFKED": "ARROW",
    "CRCW04022K00FKED": "ARROW",
    "CRCW04022K21FKED": "ARROW",
    "CRCW04022K26FKED": "מאוזר",
    "CRCW04022K37FKED": "ARROW",
    "CRCW04022K80FKED": "ARROW",
    "CRCW04022R74FKED": "אקומל ישראל בע\"מ",
    "CRCW040230K1FKED": "ARROW",
    "CRCW040230K9FKED": "ARROW",
    "CRCW040230K9FKTD": "אקומל ישראל בע\"מ",
    "CRCW040230R1FKED": "ARROW",
    "CRCW0402332KFKED": "ARROW",
    "CRCW040233K2FKED": "ARROW",
    "CRCW040233R2FKED": "ARROW",
    "CRCW04023K01FKED": "ARROW",
    "CRCW04023K32FKED": "ARROW",
    "CRCW04023K83FKED": "ARROW",
    "CRCW04023K92FKED": "ARROW",
    "CRCW040242K2FKED": "ARROW",
    "CRCW0402432RFKED": "ARROW",
    "CRCW040247R5FKED": "ARROW",
    "CRCW040249R9FKED": "ARROW",
    "CRCW04024K02FKED": "ARROW",
    "CRCW04024K22FKED": "ARROW",
    "CRCW04024K70FKED": "ARROW",
    "CRCW0402511RFKED": "ARROW",
    "CRCW040251K1FKED": "ARROW",
    "CRCW040251R1FKED": "ARROW",
    "CRCW040252K3FKED": "ARROW",
    "CRCW040256K2FKED": "ARROW",
    "CRCW04025K11FKED": "ARROW",
    "CRCW04025K49FKED": "ARROW",
    "CRCW04025K49FKTD": "אקומל ישראל בע\"מ",
    "CRCW04025K62FKED": "ARROW",
    "CRCW04025R62FKED": "אקומל ישראל בע\"מ",
    "CRCW040261K9FKTD": "אקומל ישראל בע\"מ",
    "CRCW040268R1FKED": "ARROW",
    "CRCW04026K65FKED": "ARROW",
    "CRCW040271R5FKED": "מאוזר",
    "CRCW040275R0FKED": "ARROW",
    "CRCW04027K32FKED": "ARROW",
    "CRCW04027K50FKED": "ARROW",
    "CRCW0402825RFKED": "ARROW",
    "CRCW040282R5FKED": "ARROW",
    "CRCW04028K25FKED": "ARROW",
    "CRCW04028K87FKED": "ARROW",
    "CRCW06030000Z0EA": "ARROW",
    "CRCW0603100KFKEA": "ARROW",
    "CRCW0603100RFKEA": "ARROW",
    "CRCW060310K0FKEA": "ARROW",
    "CRCW060310R0FKEA": "ARROW",
    "CRCW060311K5FKEA": "ARROW",
    "CRCW0603121KFKTA": "אקומל ישראל בע\"מ",
    "CRCW060319K1FKEA": "ARROW",
    "CRCW06031K00FKEA": "ARROW",
    "CRCW06031K00FKTA": "אקומל ישראל בע\"מ",
    "CRCW06031M00FKEA": "ARROW",
    "CRCW06031M21FKEA": "מאוזר",
    "CRCW06031R00FKEA": "אקומל ישראל בע\"מ",
    "CRCW0603221RFKTA": "אקומל ישראל בע\"מ",
    "CRCW06032K00FKEA": "ARROW",
    "CRCW06032R00FKEA": "אקומל ישראל בע\"מ",
    "CRCW060331K6FKTA": "אקומל ישראל בע\"מ",
    "CRCW060333K2FKTA": "אקומל ישראל בע\"מ",
    "CRCW060348K7FKEA": "ARROW",
    "CRCW060351K1FKEA": "ARROW",
    "CRCW06035K23FKEA": "ARROW",
    "CRCW06035K62FKEA": "ARROW",
    "CRCW060386K6FKEA": "ARROW",
    "CRCW12060000Z0EA": "ARROW",
    "CRCW12060000ZSEA": "ARROW",
    "CY8C5868AXI-LP035": "קונברג",
    "D38999/24WF35SN": "אלימק הנדסה אלקטרו-מכנית (1988) בע\"מ",
    "D38999/33W11N": "אלימק הנדסה אלקטרו-מכנית (1988) בע\"מ",
    "D38999/33W13N": "ריי-קיו בע\"מ",
    "D38999/33W17N": "אלימק הנדסה אלקטרו-מכנית (1988) בע\"מ",
    "D38999/33W19N": "ELECTRO ENTERPRISES NEW YORK LLC",
    "DAC3484IRKDT": "אלבטק תעשייה ולוגיסטיקה בע\"מ",
    "DEA205375BT-2054A1": "מאוזר",
    "DF3216-WR20KNET/LF": "פיניקס טכנולוגיות בע\"מ",
    "DFDO4125606": "פלבורג",
    "DFDO4125611": "פלבורג",
    "DFDO4125613": "פלבורג",
    "DLW21SN900SQ2L": "אס.טי.גי. אינטרנשיונל בע\"מ",
    "DP-190 GRAY 48.5ML": "ליוגב",
    "DS125BR401ANJYT": "אלבטק תעשייה ולוגיסטיקה בע\"מ",
    "DS1682S+": "ARROW",
    "DS620U+TR": "ARROW",
    "DS90LV017ATM/NOPB": "אלבטק תעשייה ולוגיסטיקה בע\"מ",
    "DS90LV031ATMTC/NOPB": "אלבטק תעשייה ולוגיסטיקה בע\"מ",
    "DS90LV032ATMTCX/NOPB": "אלבטק תעשייה ולוגיסטיקה בע\"מ",
    "DS90LV032ATMX/NOPB": "אלבטק תעשייה ולוגיסטיקה בע\"מ",
    "DS91M040TSQE/NOPB": "אלבטק תעשייה ולוגיסטיקה בע\"מ",
    "DSC612PA3A-01EQT": "EBV",
    "EC2216A/B GRAY KT2OZ": "ליוגב",
    "EI0095-109": "BFE",
    "ERF8-011-07.0-S-DV-K-TR": "SAMTEC Inc.",
    "ERF8-020-05.0-S-DV-K-TR": "SAMTEC Inc.",
    "ERM8-011-08.0-S-DV-K-TR": "SAMTEC Inc.",
    "ERM8-020-02.0-S-DV-K-TR": "SAMTEC Inc.",
    "FDN339AN": "ARROW",
    "FM24V10-G": "ARROW",
    "FT245RL": "BFE",
    "FTSH-105-01-L-DV-K-P-TR": "SAMTEC Inc.",
    "GBLC03-LF-T7": "אל-גב אלקטרוניקה",
    "GBLC05C-LF-T7": "אל-גב אלקטרוניקה",
    "GRM0335C1H1R8WA01D": "ARROW",
    "GRM0335C1H1R9WA01D": "מאוזר",
    "GRM0335C1H2R2WA01D": "מאוזר",
    "GRM0335C1H330FA01D": "אס.טי.גי. אינטרנשיונל בע\"מ",
    "GRM0335C1H3R3WA01D": "ARROW",
    "GRM0335C1HR20WA01D": "ARROW",
    "GRM0335C1HR80WA01D": "ARROW",
    "GRM0335C2A3R0WA01D": "ARROW",
    "GRM0335C2AR70WA01D": "ARROW",
    "GRM033R71E472KE14D": "ARROW",
    "GRM033Z71C104KE14D": "ARROW",
    "GRM1555C1H102GA01D": "ARROW",
    "GRM1555C1H220GA01D": "ARROW",
    "GRM1555C1H221GA01D": "אס.טי.גי. אינטרנשיונל בע\"מ",
    "GRM1555C1H3R3WA01D": "ARROW",
    "GRM1555C1H680GA01D": "ARROW",
    "GRM1555C1H7R5WA01D": "מאוזר",
    "GRM1555C1H9R1WA01D": "מאוזר",
    "GRM155C80J106ME11D": "אס.טי.גי. אינטרנשיונל בע\"מ",
    "GRM155R70J105KA12D": "ARROW",
    "GRM155R71C104KA88D": "ARROW",
    "GRM155R71H331KA01D": "ARROW",
    "GRM155R71H681KA01D": "אס.טי.גי. אינטרנשיונל בע\"מ",
    "GRM188Z71A106KA73D": "ARROW",
    "GRM21BC71H475KE11L": "ARROW",
    "GRM32EC70J107ME15L": "ARROW",
    "GRM32ER71E226KE15L": "ARROW",
    "GXO-U108L/BI-48.00MHZ": "BFE",
    "HCPL-063L-500E": "מאוזר",
    "HFCW-7500+": "MINI-CIRCUITS INT",
    "HMC1055LP2CETR": "אלבטק תעשייה ולוגיסטיקה בע\"מ",
    "HMC1096LP3E": "אלבטק תעשייה ולוגיסטיקה בע\"מ",
    "HMC424ALP3ETR": "אלבטק תעשייה ולוגיסטיקה בע\"מ",
    "HMC441LC3BTR": "אלבטק תעשייה ולוגיסטיקה בע\"מ",
    "HMC441LP3ETR": "אלבטק תעשייה ולוגיסטיקה בע\"מ",
    "HMC451LP3E": "אלבטק תעשייה ולוגיסטיקה בע\"מ",
    "HMC547ALP3E": "אלבטק תעשייה ולוגיסטיקה בע\"מ",
    "HMC564LC4TR": "אלבטק תעשייה ולוגיסטיקה בע\"מ",
    "HMC903LP3E": "אלבטק תעשייה ולוגיסטיקה בע\"מ",
    "HSMC-C120": "DIGI KEY",
    "HSSR-7110#300": "מאוזר",
    "HTST-105-01-L-DV": "SAMTEC Inc.",
    "IAC-2512WA15D0": "International Manufacturing Services Inc., (IMS)",
    "IHLP2525CZERR33M01": "ARROW",
    "INA260AIPWR": "אלבטק תעשייה ולוגיסטיקה בע\"מ",
    "IPD034N06N3G": "ARROW",
    "IRF7317TRPBF": "BFE",
    "JCM8700T10K5XM1-P-ELT": "אסקוטק אלקטרוניקה בע\"מ",
    "JCM8700T10K6XM1-P-ELT": "אסקוטק אלקטרוניקה בע\"מ",
    "JMK105BC6475MV-F": "מאוזר",
    "LMK00105SQE/NOPB": "אלבטק תעשייה ולוגיסטיקה בע\"מ",
    "LMX2592RHAT": "אלבטק תעשייה ולוגיסטיקה בע\"מ",
    "LOCTITE 222-50CC": "רוטל דבקים וכימיקלים בע\"מ",
    "LOCTITE 263-50CC": "האאס טי סי אמ 오ף ישראל אינק.",
    "LQW03AW10NJ00D": "ARROW",
    "LQW03AW5N8J00D": "ARROW",
    "LQW03AW6N8J00D": "מאוזר",
    "LQW15AN15NG80D": "ARROW",
    "LQW18AN33NJ00D": "מאוזר",
    "LQW18AN82NG00D": "ARROW",
    "LT1965EDD#PBF": "אלבטק תעשייה ולוגיסטיקה בע\"מ",
    "LTC2877HMSE#PBF": "אלבטק תעשייה ולוגיסטיקה בע\"מ",
    "LTH-010-01-G-D-A-K-TR": "SAMTEC Inc.",
    "LTM4608AIV#PBF": "אלבטק תעשייה ולוגיסטיקה בע\"מ",
    "LTM4618IV#PBF": "אלבטק תעשייה ולוגיסטיקה בע\"מ",
    "LTM4622IY#PBF": "אלבטק תעשייה ולוגיסטיקה בע\"מ",
    "LTM4634IY": "אלבטק תעשייה ולוגיסטיקה בע\"מ",
    "LTM4643MPV#PBF": "אלבטק תעשייה ולוגיסטיקה בע\"מ",
    "LTM4644IY#PBF": "אלבטק תעשייה ולוגיסטיקה בע\"מ",
    "LTM8003HY": "אלבטק תעשייה ולוגיסטיקה בע\"מ | DIGI KEY",
    "LTM8003Hy": "אלבטק תעשייה ולוגיסטיקה בע\"מ | DIGI KEY",
    "LTM8026IV#PBF": "אלבטק תעשייה ולוגיסטיקה בע\"מ",
    "M39012/25-0020": "ELECTRO ENTERPRISES NEW YORK LLC",
    "M6-10-724": "נורטק אמי בע\"מ",
    "M6-6-724": "נורטק אמי בע\"מ",
    "M80-331": "שירטק בע\"מ",
    "M83528/002B024": "ELECTRO ENTERPRISES NEW YORK LLC",
    "MAX14595ETA+T": "ARROW",
    "MAX3232EETE+": "ARROW",
    "MC74ACT04DR2G": "מאוזר",
    "MI0603J601R-10": "DIGI KEY",
    "MIC2587-1YM-TR": "ARROW",
    "MK-3C3-051-223-2400-EMI": "אס.טי.גי. אינטרנשיונל בע\"מ",
    "MM3Z3V3T1G": "ARROW",
    "MM3Z3V9T1G": "BFE",
    "MM3Z5V1T1G": "ARROW",
    "MM8430-2610RA1": "אס.טי.גי. אינטרנשיונל בע\"מ",
    "MMBD4148-7-F": "ARROW",
    "MMSZ5245BT1G": "ARROW",
    "MS15795-801": "WESCO",
    "MS15795-802": "BFE",
    "MS16995-18": "KAYTECH",
    "MS16995-2": "BFE",
    "MS16995-25": "BFE",
    "MS24693-C2": "Hardware Specialty Co., Inc.",
    "MS24693-C270": "Hardware Specialty Co., Inc.",
    "MS24693-C273": "Hardware Specialty Co., Inc.",
    "MS24693-C3": "WESCO",
    "MS3367-1-9": "WESCO",
    "MS35338-134": "Hardware Specialty Co., Inc.",
    "MS35338-135": "Hardware Specialty Co., Inc.",
    "MS35338-136": "Hardware Specialty Co., Inc.",
    "MS35338-137": "BFE",
    "MS35338-139": "WESCO",
    "MS35649-224": "Hardware Specialty Co., Inc.",
    "MS35649-244": "Hardware Specialty Co., Inc.",
    "MS51957-14": "WESCO",
    "MS51957-15": "Hardware Specialty Co., Inc.",
    "MS51957-19": "Hardware Specialty Co., Inc.",
    "MS51957-2": "WESCO",
    "MS51957-29": "WESCO",
    "MS51957-3": "WESCO",
    "MS51957-30": "WESCO",
    "MS51957-32": "WESCO",
    "MS51957-34": "WESCO",
    "MS51957-53": "HARDWARE",
    "MS51957-54": "HARDWARE",
    "MS51957-8": "WESCO",
    "MT25QL01GBBB8ESF-0AAT": "ARROW",
    "MT25QL512ABB1EW9-0SIT": "ARROW",
    "MT25QL512ABB8ESF-0AATTR": "BFE",
    "MTVA-0300N04W3S": "BFE",
    "MTVA-0500N05S": "BFE",
    "MTVA0400N05W3S": "BFE",
    "NAS1149C0432R": "WESCO",
    "NAS1291C08M": "HARDWARE",
    "NAS1291C3M": "WESCO",
    "NAS1351N4-10": "WESCO | Hardware Specialty Co., Inc.",
    "NAS1635-00-4": "KAYTECH",
    "NAS1640-0": "Hardware Specialty Co., Inc.",
    "NAS620-0": "KAYTECH",
    "NAS620C10": "BFE",
    "NAS620C2": "Hardware Specialty Co., Inc.",
    "NAS620C4": "Hardware Specialty Co., Inc.",
    "NAS620C6": "Hardware Specialty Co., Inc.",
    "NAS620C6L": "Hardware Specialty Co., Inc.",
    "NAS620C8": "BFE",
    "NAS662C2R12": "WESCO",
    "NAS662C2R4": "BFE",
    "NAS662C2R5": "WESCO",
    "NB6L239MNG": "BFE",
    "NC7ST04P5X": "ARROW",
    "NC7SZ08P5X": "ARROW",
    "NC7SZ125P5X": "מאוזר",
    "NFM18PS105R0J3D": "אס.טי.גי. אינטרנשיונל בע\"מ",
    "NL27WZ125USG-F22190": "ARROW",
    "NUP4114UPXV6T1G": "מאוזר",
    "NV24C512MUW3VTBG": "ARROW",
    "PCA9517ADP.118": "ARROW",
    "PDW06089": "אלינה - הנדסת אלקטרוניקה (2000) בע\"מ | מייקטלוג-55 יח' ל FAI",
    "PLT1.5M": "אלכסנדר שניידר",
    "PSR05-LF-T7": "אל-גב אלקטרוניקה",
    "PTL-10-724": "נורטק אמי בע\"מ",
    "PTL-9-724": "נורטק אמי בע\"מ",
    "PTN3365BSMP": "מאוזר",
    "PX-501-0059-1920M0": "אס.טי.גי. אינטרנשיונל בע\"מ",
    "REF192ESZ-REEL7": "אלבטק תעשייה ולוגיסטיקה בע\"מ",
    "SALDAVEX F 0.12MM": "BFE",
    "SALDAVEX F 0.20MM": "BFE",
    "SBB-5089Z": "BFE",
    "SEAF-20-06.0-S-06-1-AKTR": "SAMTEC Inc.",
    "SEAF-30-06-S-06-1-A-K-TR": "SAMTEC Inc.",
    "SEAF8-10-05.0-STL-04-1": "SAMTEC Inc.",
    "SEAF8-20-05.0-S-10-3": "SAMTEC Inc.",
    "SEAM-20-09.0-S-06-2-AKTR": "SAMTEC Inc.",
    "SEAM-30-03-S-06-1-A-K-TR": "SAMTEC Inc.",
    "SEAM8-10-S02.0-STL-04-1": "SAMTEC Inc.",
    "SEAM8-20-S02.0-S-10-3": "SAMTEC Inc.",
    "SFM-125-02-S-D-A-K-TR": "SAMTEC Inc.",
    "SI5513CDC-T1-GE3": "מאוזר",
    "SIA517DJ-T1-GE3": "ARROW | מאוזר",
    "SMBJ12A/TR": "DIGI KEY",
    "SMBJ36A-E3/52": "ARROW",
    "SMBJ5.0A-E3/52": "ARROW",
    "SMDA05LCC-LF": "אל-גב אלקטרוניקה",
    "SMMBT2222ALT1G": "ARROW",
    "SMMBT2907ALT1G": "ARROW",
    "SN65LVDT390PWG4": "ASTUTE",
    "SN74LVC1G07DCKT": "אלבטק תעשייה ולוגיסטיקה בע\"מ",
    "SN74LVC1G3157DRLR": "אלבטק תעשייה ולוגיסטיקה בע\"מ",
    "SN74LVC244APWT": "אלבטק תעשייה ולוגיסטיקה בע\"מ",
    "SN74LVC541APWR": "אלבטק תעשייה ולוגיסטיקה בע\"מ",
    "SN74LVCH244APW": "אלבטק תעשייה ולוגיסטיקה בע\"מ",
    "SN74LVCH8T245RHLR": "אלבטק תעשייה ולוגיסטיקה בע\"מ",
    "SN74LVT8996PW": "אלבטק תעשייה ולוגיסטיקה בע\"מ",
    "SOLC-115-02-L-Q-A-K-TR": "SAMTEC Inc.",
    "SST1-5M": "אלכסנדר שניידר",
    "SY89833LMG": "אס.טי.גי. אינטרנשיונל בע\"מ",
    "SY89833LMGTR": "אס.טי.גי. אינטרנשיונל בע\"מ",
    "T491B225K020AH": "ניסקו פרוייקטים אלקטרוניקה ותקשורת 1999 בע\"מ",
    "T491D157K016AH": "מאוזר",
    "T491X107K020AH": "ניסקו פרוייקטים אלקטרוניקה ותקשורת 1999 בע\"מ",
    "T495D107K010AHE100": "ניסקו פרוייקטים אלקטרוניקה ותקשורת 1999 בע\"מ",
    "T510X227K016ATE025": "מאוזר",
    "T521D107M016ATE050": "ARROW",
    "T521X107M025ATE060": "מייקטלוג",
    "T521X337M016ATE025": "ניסקו פרוייקטים אלקטרוניקה ותקשורת 1999 בע\"מ",
    "T521X476M035ATE030": "ARROW",
    "T528Z477M2R5ATE005": "ARROW",
    "T543O686M040ATE035": "ניסקו פרוייקטים אלקטרוניקה ותקשורת 1999 בע\"מ",
    "T55B107M6R3C0040": "ARROW | מאוזר",
    "TAJD476K020RNJ": "פיניקס טכנולוגיות בע\"מ",
    "TBMDB345-51N-Y2.5-05": "BFE",
    "TFM-125-02-S-D-A-K-TR": "SAMTEC Inc.",
    "TFS1950F": "אס.טי.גי. אינטרנשיונל בע\"מ",
    "THS4031IDGNR": "אלבטק תעשייה ולוגיסטיקה בע\"מ",
    "THT-1-724-10": "נורטק אמי בע\"מ",
    "TM1S4-C": "אלכסנדר שניידר",
    "TOLC-115-22-S-Q-A-K-TR": "SAMTEC Inc.",
    "TPS2002CDRCT": "אלבטק תעשייה ולוגיסטיקה בע\"מ",
    "TPS22810DRVT": "אלבטק תעשייה ולוגיסטיקה בע\"מ",
    "TPS3808G01DRVT": "אלבטק תעשייה ולוגיסטיקה בע\"מ",
    "TPS72301DBVR": "אלבטק תעשייה ולוגיסטיקה בע\"מ",
    "TPS7A8300RGRT": "אלבטק תעשייה ולוגיסטיקה בע\"מ",
    "TPS82085SILT": "אלבטק תעשייה ולוגיסטיקה בע\"מ",
    "TR3C107K010C0100": "אקומל ישראל בע\"מ",
    "TRF37A73IDSGT": "אלבטק תעשייה ולוגיסטיקה בע\"מ",
    "TRM-2090-MC-SMA-02": "פי אי איי - ג'נסיס ישראל בע\"מ",
    "TS0500S": "BFE",
    "TS0501W3S/TR": "BFE",
    "TS0505S": "BFE",
    "TUSB2046BIRHBT": "אלבטק תעשייה ולוגיסטיקה בע\"מ",
    "TV07RW17-73S(S25AD)": "BFE",
    "TXB0304RUTR": "אלבטק תעשייה ולוגיסטיקה בע\"מ",
    "UPS-06-07.0-03-L-PV": "SAMTEC Inc.",
    "UPT-06-03.0-03-L-PV": "SAMTEC Inc.",
    "VFXO321-DGED-125MHZ": "דונטק אלקטרוניקה בע\"מ",
    "W25Q128JVSIQ": "ARROW",
    "WP4P+": "MINI-CIRCUITS INT",
    "WP4S+": "MINI-CIRCUITS INT",
    "WSL25121L000FEA": "ARROW",
    "WTBV36PD11SY": "אס.טי.גי. אינטרנשיונל בע\"מ",
}


st.set_page_config(
    page_title="MRP Executive Control Tower",
    page_icon="🚀",
    layout="wide"
)

# ==========================================================
# AUTHENTICATION GATE (שכבת הגנה למערכת)
# ==========================================================
def check_password():
    def password_entered():
        if st.secrets["passwords"].get(st.session_state["username"]) == st.session_state["password"]:
            st.session_state["password_correct"] = True
            del st.session_state["password"]
            del st.session_state["username"]
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.markdown("### Designed By Oren Amram and Yossi Amar🔐 כניסה למערכת  בקרת חוסרים (MRP)")
        st.text_input("שם משתמש", key="username")
        st.text_input("סיסמה", type="password", key="password")
        st.button("התחבר", on_click=password_entered)
        return False
    elif not st.session_state["password_correct"]:
        st.markdown("### Designed By Oren Amram and Yossi Amar🔐 כניסה למערכת  בקרת חוסרים (MRP)")
        st.text_input("שם משתמש", key="username")
        st.text_input("סיסמה", type="password", key="password")
        st.button("התחבר", on_click=password_entered)
        st.error("😕 שם משתמש או סיסמה שגויים")
        return False
    else:
        return True

if not check_password():
    st.stop() 

# ==========================================================
# SUPABASE SETUP - מוקדם בקובץ בכוונה, כי גם בורר הפרויקטים (למטה) צריך
# אותו כדי לטעון/לשמור פרויקטים שנוספו דרך "קישור מותאם אישית".
# ==========================================================
SUPABASE_URL = st.secrets["supabase"]["url"]
SUPABASE_KEY = st.secrets["supabase"]["key"]

@st.cache_resource
def init_supabase() -> Client:
    return create_client(SUPABASE_URL, SUPABASE_KEY)

supabase = init_supabase()

@st.cache_data(ttl=60)
def fetch_saved_projects():
    """פרויקטים שנשמרו דרך 'קישור מותאם אישית' + כפתור השמירה - מצטרפים
    אוטומטית לרשימה הקבועה AVAILABLE_PROJECTS, בלי לגעת בקוד."""
    try:
        response = supabase.table("mrp_projects").select("*").execute()
        if response.data:
            return {row["project_id"]: row["url"] for row in response.data}
    except Exception:
        pass
    return {}

def save_project_to_cloud(project_id, url):
    try:
        supabase.table("mrp_projects").upsert(
            {"project_id": project_id, "url": url}, on_conflict="project_id"
        ).execute()
        fetch_saved_projects.clear()
        return True, None
    except Exception as e:
        return False, str(e)

def delete_project_from_cloud(project_id):
    try:
        supabase.table("mrp_projects").delete().eq("project_id", project_id).execute()
        fetch_saved_projects.clear()
        return True, None
    except Exception as e:
        return False, str(e)

# ==========================================================
# בחירת פרויקט - עובדים עם כמה קבצי Excel (פרויקטים) במקביל, בלי לגעת בקוד.
# יש שתי דרכים להוסיף פרויקט קבוע: (א) להוסיף שורה למילון AVAILABLE_PROJECTS
# למטה בקוד, או (ב) להזין "קישור מותאם אישית" ולחיצה על "💾 שמור פרויקט זה
# לקבע" - זה נשמר בענן (Supabase) ומצטרף אוטומטית לרשימה בכל כניסה הבאה,
# בלי לגעת בקוד בכלל. אין הגבלה על מספר הפרויקטים.
#
# חשוב: ה-project_id (המזהה שלפיו נשמרים/מסוננים מלאי, WIP ותוכנית הרכבה)
# הוא לא נגזר מתוך שם הקובץ באופן אוטומטי - כי אם אתה מעלה כל חודש קובץ עם
# שם חדש (למשל עם התאריך בשם), גזירה אוטומטית משם הקובץ הייתה יוצרת בטעות
# "פרויקט" חדש בכל פעם ומאבדת את הרצף. במקום זה, אתה קובע את שם הפרויקט
# בעצמך פעם אחת - וכל עוד אתה משתמש באותו שם, זה תמיד אותו project_id, גם
# אם שם הקובץ עצמו משתנה.
# ==========================================================
AVAILABLE_PROJECTS_HARDCODED = {
    "אנטנה (ברירת מחדל)": GITHUB_URL,
    "WISLAB": "https://raw.githubusercontent.com/orenamram-arch/offset/main/NEW%20PROJECT%20WISLAB%20JUNE%202026%20AVL.xlsx",
}

def _normalize_github_url(url):
    """אם הודבק קישור רגיל לצפייה בגיטהאב (עם github.com/.../blob/...) במקום קישור
    raw, ממירים אותו אוטומטית - כדי שלא יהיה צורך לחפש את כפתור ה-'Raw' בכלל."""
    url = (url or "").strip()
    if "github.com" in url and "raw.githubusercontent.com" not in url and "/blob/" in url:
        url = url.replace("github.com", "raw.githubusercontent.com").replace("/blob/", "/")
    return url

# פרויקטים קבועים בקוד + פרויקטים ששמרת בעבר דרך כפתור "שמור פרויקט זה לקבע"
# (השמורים-בענן גוברים אם יש התנגשות שם, כי הם העדכניים ביותר)
AVAILABLE_PROJECTS = {**AVAILABLE_PROJECTS_HARDCODED, **fetch_saved_projects()}

with st.sidebar:
    st.markdown("### 📁 פרויקט פעיל")
    _project_names = list(AVAILABLE_PROJECTS.keys()) + ["🔗 קישור מותאם אישית..."]
    # אם הגעת לאפליקציה עם ?project=<שם> בכתובת, נבחר אותו כברירת מחדל -
    # כך אפשר לשמור/לשלוח קישור שתמיד קופץ ישר לפרויקט הנכון.
    _default_idx = 0
    try:
        _qp_project = st.query_params.get("project")
        if _qp_project in _project_names:
            _default_idx = _project_names.index(_qp_project)
    except Exception:
        pass
    _selected_project = st.selectbox("בחר פרויקט", _project_names, index=_default_idx, key="selected_project_name")
    try:
        st.query_params["project"] = _selected_project
    except Exception:
        pass

    if _selected_project == "🔗 קישור מותאם אישית...":
        PROJECT_ID = st.text_input(
            "שם הפרויקט (קבוע - זה מה שמפריד את המלאי/WIP/תוכנית בין פרויקטים)",
            value=st.session_state.get("custom_project_id_value", ""),
            key="custom_project_id_value",
            help="בחר שם יציב ותשתמש בו תמיד לפרויקט הזה, גם אם שם קובץ ה-Excel עצמו משתנה מדי חודש."
        )
        _raw_custom_url = st.text_input(
            "כתובת לקובץ ה-Excel (אפשר להדביק גם קישור רגיל לצפייה בגיטהאב - הוא יומר אוטומטית)",
            value=st.session_state.get("custom_project_url_value", ""),
            key="custom_project_url_value"
        )
        ACTIVE_PROJECT_URL = _normalize_github_url(_raw_custom_url)
        if not ACTIVE_PROJECT_URL or not PROJECT_ID:
            st.info("הזן שם פרויקט וכתובת לקובץ Excel כדי להמשיך.")
            st.stop()
        if PROJECT_ID in AVAILABLE_PROJECTS_HARDCODED:
            st.warning(f"השם '{PROJECT_ID}' כבר קיים ברשימה הקבועה בקוד - בחר שם אחר כדי לא להתנגש איתו.")
        elif st.button("💾 שמור פרויקט זה לקבע (יופיע ברשימה בפעם הבאה)"):
            _ok, _err = save_project_to_cloud(PROJECT_ID, ACTIVE_PROJECT_URL)
            if _ok:
                st.success(f"נשמר! '{PROJECT_ID}' יופיע ברשימה בכניסה הבאה.")
                st.rerun()
            else:
                st.error(f"שגיאה בשמירת הפרויקט: {_err}")
    else:
        ACTIVE_PROJECT_URL = _normalize_github_url(AVAILABLE_PROJECTS[_selected_project])
        PROJECT_ID = _selected_project
        if _selected_project not in AVAILABLE_PROJECTS_HARDCODED:
            if st.button("🗑️ הסר פרויקט זה מהרשימה השמורה"):
                _ok, _err = delete_project_from_cloud(_selected_project)
                if _ok:
                    st.success("הוסר.")
                    st.rerun()
                else:
                    st.error(f"שגיאה בהסרה: {_err}")

# אם עברו לפרויקט אחר מאז הטעינה הקודמת, מנקים נתוני-session שתלויים בפרויקט
# הספציפי (כמו תוכנית הרכבה), כדי שלא יישארו נתונים "ישנים" מפרויקט קודם.
if st.session_state.get("_loaded_project_id") != PROJECT_ID:
    st.session_state.pop("custom_assembly_plan_df", None)
    st.session_state["_loaded_project_id"] = PROJECT_ID

# ==========================================================
# GLOBAL THEME / CSS
# ==========================================================
PRIMARY = "#4F46E5"
PRIMARY_DARK = "#3730A3"
ACCENT = "#06B6D4"
DANGER = "#EF4444"
WARNING = "#F59E0B"
SUCCESS = "#10B981"

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Assistant:wght@400;600;700;800&display=swap');

[data-testid="stAppViewContainer"] .main .block-container,
[data-testid="stSidebarContent"] {{
    font-family: 'Assistant', sans-serif;
    direction: rtl;
}}
[data-testid="stAppViewContainer"] .main .block-container * ,
[data-testid="stSidebarContent"] * {{
    font-family: 'Assistant', sans-serif;
}}

#MainMenu {{visibility: hidden;}}
footer {{visibility: hidden;}}

/* --- UI FIX: הסליידר משמאל לימין אך הטקסט מימין לשמאל --- */
[data-testid="stSlider"] {{
    direction: ltr !important;
}}
[data-testid="stSlider"] [data-testid="stMarkdownContainer"] {{
    direction: rtl !important;
    text-align: right !important;
}}
/* -------------------------------------------------------- */

.hero-banner {{
    background: linear-gradient(120deg, {PRIMARY} 0%, {PRIMARY_DARK} 45%, {ACCENT} 100%);
    padding: 28px 32px;
    border-radius: 18px;
    margin-bottom: 22px;
    box-shadow: 0 10px 30px rgba(79,70,229,0.35);
}}
.hero-banner h1 {{
    color: white;
    font-weight: 800;
    font-size: 30px;
    margin: 0;
}}
.hero-banner p {{
    color: rgba(255,255,255,0.9);
    font-size: 15px;
    margin-top: 6px;
}}

.kpi-card, .kanban-card {{
    background-color: var(--secondary-background-color, #ffffff);
    color: var(--text-color, #111827);
    border: 1px solid rgba(128,128,128,0.25);
    border-radius: 14px;
    padding: 18px 16px;
    text-align: center;
    box-shadow: 0 4px 14px rgba(0,0,0,0.1);
    transition: transform 0.15s ease;
}}
.kpi-card:hover {{ transform: translateY(-3px); }}
.kpi-label {{
    font-size: 13px;
    opacity: 0.75;
    margin-bottom: 6px;
    font-weight: 600;
}}
.kpi-value {{
    font-size: 30px;
    font-weight: 800;
}}
.kpi-sub {{
    font-size: 12px;
    opacity: 0.6;
    margin-top: 4px;
}}

.kpi-green {{ border-top: 4px solid {SUCCESS}; }}
.kpi-red {{ border-top: 4px solid {DANGER}; }}
.kpi-orange {{ border-top: 4px solid {WARNING}; }}
.kpi-blue {{ border-top: 4px solid {ACCENT}; }}

.section-title {{
    font-weight: 800;
    font-size: 19px;
    margin: 18px 0 10px 0;
    border-right: 4px solid {PRIMARY};
    padding-right: 10px;
    color: var(--text-color, inherit);
}}

.kanban-col-header {{
    font-weight: 800;
    font-size: 15px;
    padding: 8px 12px;
    border-radius: 10px;
    margin-bottom: 8px;
    text-align: center;
}}
.kanban-card {{
    border: 1px solid rgba(128,128,128,0.2);
    border-radius: 10px;
    padding: 10px 12px;
    margin-bottom: 8px;
    border-right: 3px solid {PRIMARY};
    font-size: 13px;
}}

.exec-summary-strip {{
    display: flex;
    flex-wrap: wrap;
    gap: 14px;
    background: linear-gradient(135deg, rgba(79,70,229,0.10) 0%, rgba(6,182,212,0.10) 100%);
    border: 1px solid rgba(79,70,229,0.25);
    border-radius: 16px;
    padding: 16px 20px;
    margin-bottom: 18px;
    position: sticky;
    top: 0;
    z-index: 999;
    backdrop-filter: blur(6px);
}}
.exec-stat {{
    flex: 1;
    min-width: 150px;
    text-align: center;
    padding: 6px 10px;
    border-right: 1px solid rgba(128,128,128,0.2);
}}
.exec-stat:last-child {{ border-right: none; }}
.exec-stat-icon {{ font-size: 22px; margin-bottom: 2px; }}
.exec-stat-value {{
    font-size: 26px;
    font-weight: 800;
    background: linear-gradient(90deg, {PRIMARY}, {ACCENT});
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    background-clip: text;
}}
.exec-stat-label {{ font-size: 12px; opacity: 0.75; font-weight: 600; margin-top: 2px; }}

.nav-bar-label {{
    font-weight: 700;
    font-size: 13px;
    opacity: 0.7;
    margin-bottom: 6px;
}}
div[data-testid="stPopover"] > button {{
    border-radius: 12px !important;
    font-weight: 700 !important;
    padding: 10px 14px !important;
    border: 1.5px solid rgba(79,70,229,0.3) !important;
    background: linear-gradient(135deg, rgba(79,70,229,0.10), rgba(6,182,212,0.10)) !important;
    transition: all 0.15s ease;
}}
div[data-testid="stPopover"] > button:hover {{
    background: linear-gradient(135deg, {PRIMARY}, {ACCENT}) !important;
    color: white !important;
    transform: translateY(-2px);
    box-shadow: 0 6px 16px rgba(79,70,229,0.35);
}}
[data-testid="stPopoverBody"] button[kind="secondary"],
[data-testid="stPopoverBody"] button[kind="primary"] {{
    text-align: right !important;
    justify-content: flex-start !important;
    border-radius: 8px !important;
    margin-bottom: 3px !important;
    font-weight: 600 !important;
}}
[data-testid="stPopoverBody"] button[kind="primary"] {{
    background: linear-gradient(135deg, {PRIMARY}, {PRIMARY_DARK}) !important;
    color: white !important;
    border: none !important;
}}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="hero-banner">
    <h1>🚀 MRP Executive Control Tower & Decision Hub</h1>
    <p>מערכת ניהול חוסרים מתקדמת, סימולציות קבלת החלטות (What-If), ותמונת מצב ניהולית מסונכרנת לענן במהירות שיא</p>
</div>
""", unsafe_allow_html=True)

def kpi_card(label, value, sub="", color="blue"):
    st.markdown(f"""
    <div class="kpi-card kpi-{color}">
        <div class="kpi-label">{label}</div>
        <div class="kpi-value">{value}</div>
        <div class="kpi-sub">{sub}</div>
    </div>
    """, unsafe_allow_html=True)

try:
    _theme_base = st.get_option("theme.base")
except Exception:
    _theme_base = None

PLOTLY_TEMPLATE = "plotly_white" if _theme_base == "light" else "plotly_dark"
COLOR_SEQ = [PRIMARY, ACCENT, WARNING, DANGER, SUCCESS, "#A78BFA", "#F472B6", "#34D399"]

@st.cache_data(ttl=60)
def fetch_all_inventory_records(project_id):
    try:
        response = supabase.table("mrp_inventory_updates").select("*").eq("project_id", project_id).execute()
        records = {}
        if response.data:
            for row in response.data:
                pn = str(row.get("pn")).strip()
                eta_val = row.get("eta", "")
                if not eta_val or str(eta_val).strip() in ["", "None", "NaT", "nan"]:
                    eta_val = ""
                status_val = row.get("status", "פתוח") or "פתוח"
                records[pn] = {
                    "added_stock": float(row.get("added_stock", 0.0) or 0.0),
                    "eta": eta_val,
                    "status": status_val,
                    "supplier": row.get("supplier", "") or "",
                    "comment": row.get("comment", ""),
                    "updated_by": row.get("updated_by", ""),
                    "updated_at": row.get("updated_at", ""),
                    "item_type": row.get("item_type", "") or "",
                    "unit_price": row.get("unit_price", None)
                }
        return records
    except Exception:
        return {}

def get_inventory_record(pn, cache=None):
    all_recs = cache if cache is not None else fetch_all_inventory_records(PROJECT_ID)
    res = all_recs.get(str(pn).strip())
    if res:
        return (
            res["added_stock"],
            res["eta"],
            res["status"],
            res["supplier"],
            res["comment"],
            res["updated_by"],
            res["updated_at"]
        )
    return 0.0, "", "פתוח", "", "", "", ""

def get_effective_item_type(pn, original_type, cache=None):
    all_recs = cache if cache is not None else fetch_all_inventory_records(PROJECT_ID)
    res = all_recs.get(str(pn).strip())
    if res:
        override = res.get("item_type", "")
        if override and str(override).strip() not in ["", "None", "nan"]:
            return str(override)
    return original_type

def get_effective_supplier(pn, cache=None):
    """מחזיר את הספק בפועל לפריט: קודם שינוי ידני שנשמר, ואם אין - הספק בפועל
    מהקובץ (עמודת AW_COL, שנמצאת דינמית לפי שם - 'ספק'/'SUPPLIER' וכו').
    לא מחזיר אף פעם ברירת מחדל קבועה שלא בהכרח נכונה לפריט הספציפי."""
    all_recs = cache if cache is not None else fetch_all_inventory_records(PROJECT_ID)
    res = all_recs.get(str(pn).strip())
    if res:
        override = res.get("supplier", "")
        if override and str(override).strip() not in ["", "None", "nan"]:
            return str(override)
    try:
        match = df[df[PN_COL].astype(str).str.strip() == str(pn).strip()]
        if not match.empty:
            file_supplier = str(match.iloc[0].get(AW_COL, "")).strip()
            if file_supplier and file_supplier not in ["", "None", "nan"]:
                return file_supplier
    except Exception:
        pass
    return "לא צוין"

def get_effective_price(pn, file_price, cache=None):
    """מחזיר את מחיר היחידה בפועל לפריט: אם המשתמש הזין ידנית מחיר שמור
    לפריט הזה (בפרויקט הנוכחי - כל פרויקט נשמר בנפרד), הוא קובע. אחרת - מחזיר
    את המחיר מהקובץ (file_price) כמו שהוא, ללא שינוי."""
    all_recs = cache if cache is not None else fetch_all_inventory_records(PROJECT_ID)
    res = all_recs.get(str(pn).strip())
    if res:
        override = res.get("unit_price")
        if override is not None and str(override).strip() not in ["", "None", "nan"]:
            try:
                return float(override)
            except (TypeError, ValueError):
                pass
    return file_price

@st.cache_data(ttl=60)
def fetch_wip_records(project_id):
    try:
        response = supabase.table("mrp_wip_assemblies").select("*").eq("project_id", project_id).execute()
        if response.data:
            return {str(row.get("assembly_pn")).strip(): float(row.get("wip_qty", 0.0)) for row in response.data}
    except:
        pass
    return {}

def save_wip_record(assembly_pn, wip_qty):
    current_wip_dict = fetch_wip_records(PROJECT_ID)
    existing_qty = current_wip_dict.get(str(assembly_pn), 0.0)
    total_new_qty = existing_qty + float(wip_qty)

    now_str = (datetime.utcnow() + timedelta(hours=3)).strftime("%Y-%m-%d %H:%M")
    payload = {
        "assembly_pn": str(assembly_pn),
        "wip_qty": float(total_new_qty),
        "updated_at": now_str,
        "project_id": PROJECT_ID
    }
    try:
        supabase.table("mrp_wip_assemblies").upsert(payload, on_conflict="assembly_pn,project_id").execute()
        fetch_wip_records.clear()
    except Exception as e:
        st.error(f"שגיאה בשמירת WIP ל-Supabase: {e}")

def delete_wip_record(assembly_pn):
    try:
        supabase.table("mrp_wip_assemblies").delete().eq("assembly_pn", str(assembly_pn)).eq("project_id", PROJECT_ID).execute()
        fetch_wip_records.clear()
    except Exception as e:
        st.error(f"שגיאה במחיקת WIP מ-Supabase: {e}")

def save_inventory_record(pn, added_stock, eta, status, supplier, comment, updated_by, webhook_url="", item_type=None, unit_price=None):
    now_str = (datetime.utcnow() + timedelta(hours=3)).strftime("%Y-%m-%d %H:%M")
    existing_rec = fetch_all_inventory_records(PROJECT_ID).get(str(pn).strip(), {})
    if item_type is None:
        item_type_val = existing_rec.get("item_type", "")
    else:
        item_type_val = item_type
    if unit_price is None:
        unit_price_val = existing_rec.get("unit_price", None)
    else:
        unit_price_val = unit_price
    payload = {
        "pn": str(pn),
        "added_stock": float(added_stock),
        "eta": str(eta),
        "status": str(status),
        "supplier": str(supplier),
        "comment": str(comment),
        "updated_by": str(updated_by),
        "updated_at": now_str,
        "item_type": str(item_type_val),
        "unit_price": float(unit_price_val) if unit_price_val not in [None, ""] else None,
        "project_id": PROJECT_ID
    }
    try:
        supabase.table("mrp_inventory_updates").upsert(payload, on_conflict="pn,project_id").execute()
        supabase.table("mrp_inventory_history").insert(payload).execute()
        fetch_all_inventory_records.clear()
    except Exception as e:
        st.error(f"שגיאה בשמירה ל-Supabase: {e}")

    if webhook_url:
        msg = "🔔 עדכון מלאי/ETA למוצר!\nמק'ט: " + str(pn) + "\nתוספת מלאי: " + str(added_stock) + "\nסטטוס: " + str(status) + "\nETA: " + str(eta)
        try:
            requests.post(webhook_url, data=json.dumps({"text": msg}), headers={'Content-Type': 'application/json'})
        except:
            pass

def bulk_update_suppliers(supplier_map, inv_cache=None):
    if inv_cache is None:
        inv_cache = fetch_all_inventory_records(PROJECT_ID)
    now_str = (datetime.utcnow() + timedelta(hours=3)).strftime("%Y-%m-%d %H:%M")
    payloads = []
    for pn, supplier in supplier_map.items():
        existing = inv_cache.get(pn, {})
        payloads.append({
            "pn": str(pn),
            "added_stock": float(existing.get("added_stock", 0.0) or 0.0),
            "eta": str(existing.get("eta", "") or ""),
            "status": str(existing.get("status", "פתוח") or "פתוח"),
            "supplier": str(supplier),
            "comment": str(existing.get("comment", "") or ""),
            "updated_by": "Bulk Supplier Import",
            "updated_at": now_str,
            "project_id": PROJECT_ID
        })
    try:
        supabase.table("mrp_inventory_updates").upsert(payloads, on_conflict="pn,project_id").execute()
        supabase.table("mrp_inventory_history").insert(payloads).execute()
        fetch_all_inventory_records.clear()
        return len(payloads), None
    except Exception as e:
        return 0, str(e)

def delete_inventory_record(pn):
    try:
        supabase.table("mrp_inventory_updates").delete().eq("pn", str(pn)).eq("project_id", PROJECT_ID).execute()
        fetch_all_inventory_records.clear()
    except Exception as e:
        st.error(f"שגיאה במחיקה מ-Supabase: {e}")

@st.cache_data(ttl=60)
def fetch_cloud_assembly_plan(project_id):
    try:
        response = supabase.table("mrp_assembly_plans").select("*").eq("project_id", project_id).execute()
        if response.data:
            return pd.DataFrame(response.data)
    except:
        pass
    return pd.DataFrame()

def save_cloud_assembly_plan(plan_df, project_id):
    try:
        records = plan_df.to_dict(orient="records")
        for rec in records:
            rec["project_id"] = project_id
        # מוחקים רק את השורות של הפרויקט הנוכחי - לא כל התוכן בטבלה (זה היה מוחק
        # בטעות גם תוכניות שמורות של פרויקטים אחרים).
        supabase.table("mrp_assembly_plans").delete().eq("project_id", project_id).execute()
        if records:
            supabase.table("mrp_assembly_plans").upsert(records).execute()
        fetch_cloud_assembly_plan.clear()
    except Exception as e:
        st.error(f"שגיאה בשמירת תוכנית הייצור לענן: {e}")

# ==========================================================
# DATA LOADING FROM GITHUB & SESSION STATE
# ==========================================================
@st.cache_data
def load_data(url):
    df_raw = pd.read_excel(url, header=None)
    header_row = _find_matrix_header_row(df_raw)
    if header_row is None:
        # נפילה בטוחה למבנה הידוע היסטורית, אם שורת הכותרת הראשית לא אותרה
        header_row = 29
    df = pd.read_excel(url, header=header_row)
    df.columns = [str(c).strip() if pd.notnull(c) else c for c in df.columns]
    return df, df_raw, header_row

# ==========================================================
# פונקציות גילוי מבנה כלליות - מאתרות עמודות/גבולות לפי הכותרות בפועל
# בקובץ (טקסט), ולא לפי מיקום קבוע. זה מה שמאפשר לעבוד עם קבצי פרויקטים
# שונים לגמרי במבנה (מספר הרכבות, מספר חודשים, סדר עמודות מטא-דאטה וכו').
# ==========================================================
def _find_matrix_header_row(raw_df):
    """מאתר את שורת הכותרת הראשית של הגיליון: עמודה עם '#' ואיפשהו בסמוך לה (עד 5
    עמודות קדימה - כדי לסבול עמודה ריקה ביניהן, כמו שקורה בחלק מהקבצים) עמודת
    'PN_ID'."""
    for r in range(raw_df.shape[0]):
        hash_cols = [c for c in range(min(10, raw_df.shape[1]))
                     if isinstance(raw_df.iat[r, c], str) and raw_df.iat[r, c].strip() == "#"]
        if not hash_cols:
            continue
        hash_col = hash_cols[0]
        for c2 in range(hash_col, min(hash_col + 6, raw_df.shape[1])):
            v = raw_df.iat[r, c2]
            if isinstance(v, str) and v.strip().upper() == "PN_ID":
                return r
    return None

def _find_text_cell(raw_df, text):
    """מחפש בכל הגיליון תא עם טקסט מדויק (לא תלוי-רישיות), מחזיר (row, col) ראשון שנמצא."""
    target = text.strip().upper()
    max_r, max_c = raw_df.shape
    for r in range(max_r):
        for c in range(max_c):
            v = raw_df.iat[r, c]
            if isinstance(v, str) and v.strip().upper() == target:
                return (r, c)
    return None

def _find_desc_level_pn_cols(raw_df):
    """מאתר בכל הגיליון שלישיית עמודות סמוכות עם הכותרות 'DESC'/'DESCRIPTION','LEVEL','PN'
    (סובלני גם לאותיות קטנות ולמילה המלאה 'DESCRIPTION') - זהו בלוק עץ ה-BOM המפורק.
    גם מזהה אם יש עמודת כמות ('QTY...') מיד אחרי PN."""
    max_r, max_c = raw_df.shape
    for r in range(max_r):
        for c in range(max_c - 2):
            v0, v1, v2 = raw_df.iat[r, c], raw_df.iat[r, c + 1], raw_df.iat[r, c + 2]
            if (isinstance(v0, str) and v0.strip().upper() in ("DESC", "DESCRIPTION") and
                    isinstance(v1, str) and v1.strip().upper() == "LEVEL" and
                    isinstance(v2, str) and v2.strip().upper() == "PN"):
                qty_col = None
                if c + 3 < max_c:
                    v3 = raw_df.iat[r, c + 3]
                    if isinstance(v3, str) and "QTY" in v3.strip().upper():
                        qty_col = c + 3
                return {"desc": c, "level": c + 1, "pn": c + 2, "qty": qty_col}
    return None

def _is_date_col(v):
    try:
        return pd.notnull(pd.to_datetime(v))
    except Exception:
        return False

def _longest_date_run_ending_before(named_cols, end_idx_exclusive, lookback=150):
    """מוצא את הרצף הרציף הארוך ביותר של עמודות-תאריך, בחלון שמסתיים ממש לפני end_idx_exclusive."""
    start_window = max(0, end_idx_exclusive - lookback)
    best = (0, 0, 0)
    cur_start = None
    for i in range(start_window, end_idx_exclusive):
        if _is_date_col(named_cols[i]):
            if cur_start is None:
                cur_start = i
        else:
            if cur_start is not None:
                length = i - cur_start
                if length > best[0]:
                    best = (length, cur_start, i)
                cur_start = None
    if cur_start is not None:
        length = end_idx_exclusive - cur_start
        if length > best[0]:
            best = (length, cur_start, end_idx_exclusive)
    return (best[1], best[2]) if best[0] > 0 else None

def _find_column_by_candidates(named_cols, candidates):
    """מאתר עמודה לפי רשימת שמות אפשריים (בדיקה מדויקת, לא תלוית-רישיות/רווחים)."""
    norm = {str(c).strip().upper(): c for c in named_cols}
    for cand in candidates:
        key = cand.strip().upper()
        if key in norm:
            return norm[key]
    return None

def _find_matrix_level_desc_rows(raw_df, header_row):
    """אסטרטגיה שנייה לזיהוי עץ ה-BOM: לקבצים בלי בלוק DESC/LEVEL/PN נפרד, לפעמים
    יש מעל שורת הכותרת הראשית שתי שורות-תיוג - 'DESC' ומשהו-שמכיל-'LEVEL' (סובלני
    גם לשגיאות כתיב כמו 'LEVAL') - שמתארות ישירות את עמודות מטריצת ה-where-used."""
    desc_row, level_row = None, None
    top = max(0, header_row - 15)
    for r in range(top, header_row):
        for c in range(raw_df.shape[1]):
            v = raw_df.iat[r, c]
            if isinstance(v, str):
                vv = v.strip().upper()
                if ("DESC" in vv) and desc_row is None:
                    desc_row = r
                elif ("LEVEL" in vv or "LEVAL" in vv) and level_row is None:
                    level_row = r
    return desc_row, level_row

def _valid_level_value(v):
    try:
        return int(v) >= 0
    except (ValueError, TypeError):
        return False

def _find_matrix_col_run(raw_df, level_row):
    """מוצא רצף רציף של עמודות עם ערכי-רמה תקינים בשורת ה-level, החל מהעמודה
    הראשונה שיש בה ערך כזה. עוצרים בפער הראשון - כדי לא להיתפס לערכים מספריים
    מקריים רחוק יותר בגיליון (כמו עמודת SUM, תאריכים, וכו')."""
    start = None
    for c in range(raw_df.shape[1]):
        if _valid_level_value(raw_df.iat[level_row, c]):
            start = c
            break
    if start is None:
        return []
    cols, c = [], start
    while c < raw_df.shape[1] and _valid_level_value(raw_df.iat[level_row, c]):
        cols.append(c)
        c += 1
    return cols

try:
    with st.spinner('טוען נתוני MRP מ-GitHub...'):
        df, df_raw, _header_row = load_data(ACTIVE_PROJECT_URL)
except Exception as e:
    _err_text = str(e)
    if "engine manually" in _err_text or "format cannot be determined" in _err_text:
        st.error(
            "שגיאה בטעינת הקובץ מ-GitHub: הכתובת לא מחזירה קובץ Excel בינארי תקין "
            "(פנדס לא הצליח לזהות את הפורמט). הסיבות הנפוצות ביותר:\n\n"
            "1. זה קישור רגיל לצפייה בגיטהאב (github.com/.../blob/...) ולא קישור raw "
            "(raw.githubusercontent.com/...). היכנס לקובץ בגיטהאב, לחץ על הכפתור 'Raw', והעתק את הכתובת משם.\n"
            "2. הריפו פרטי - קישורי raw לא עובדים בלי אימות בריפו פרטי.\n"
            "3. הקובץ מנוהל דרך Git LFS - raw.githubusercontent.com מחזיר אז קובץ-מצביע קטן במקום התוכן עצמו.\n\n"
            f"פירוט השגיאה המקורית: {_err_text}"
        )
    else:
        st.error(f"שגיאה בטעינת הקובץ מ-GitHub. פירוט השגיאה: {_err_text}")
    st.stop()

def _find_price_column(named_cols):
    """מאתר עמודת מחיר-ליחידה. קודם מנסה שמות מפורשים (עברית/אנגלית), ואם לא
    נמצא - נופל לחיפוש רחב יותר של כל עמודה שמכילה 'מחיר' או 'PRICE'. בשני
    המקרים תמיד מדלגים במפורש על עמודות של מחיר *כולל* (TOTAL PRICE / מחיר
    כולל) - כי אלה כבר מוכפלות בכמות-לקשר ואינן מחיר ליחידה בודדת."""
    exclude_terms = ["TOTAL", "כולל", "FINAL", "סופי"]

    def _is_total_price(col_text_upper, col_text_raw):
        return any(term in col_text_upper or term in col_text_raw for term in exclude_terms)

    explicit_candidates = ["PRICE_CALC", "UNIT PRICE", "מחיר יחידה", "TARGET PRICE", "PRICE", "מחיר"]
    norm = {str(c).strip().upper(): c for c in named_cols}
    for cand in explicit_candidates:
        key = cand.strip().upper()
        if key in norm:
            col = norm[key]
            if not _is_total_price(str(col).strip().upper(), str(col).strip()):
                return col

    for c in named_cols:
        c_raw = str(c).strip()
        c_upper = c_raw.upper()
        if ("PRICE" in c_upper or "מחיר" in c_raw) and not _is_total_price(c_upper, c_raw):
            return c
    return None

# --- עמודות מטא-דאטה בסיסיות - מאותרות לפי שם הכותרת, לא לפי מיקום קבוע ---
PN_COL = _find_column_by_candidates(df.columns, ["PN_ID"]) or (df.columns[1] if len(df.columns) > 1 else df.columns[0])
DESC_COL = _find_column_by_candidates(df.columns, ["DESCRIPTION"]) or (df.columns[4] if len(df.columns) > 4 else df.columns[-1])
ITEM_TYPE_COL = _find_column_by_candidates(df.columns, ["סיווג פריט", "ITEM_TYPE", "ITEM TYPE"])
PRICE_COL = _find_price_column(df.columns)
def _find_supplier_column(named_cols):
    """מאתר עמודת ספק. קודם שמות מפורשים, ואם לא נמצא - כל עמודה שמכילה 'ספק' או
    'SUPPLIER' בכל צירוף (עברית/אנגלית, כולל כשהם חלק משם עמודה ארוך יותר)."""
    explicit_candidates = ["ספק", "SUPPLIER", "MANUFACTURER", "POC SUPPLIER"]
    hit = _find_column_by_candidates(named_cols, explicit_candidates)
    if hit is not None:
        return hit
    for c in named_cols:
        c_raw = str(c).strip()
        if "SUPPLIER" in c_raw.upper() or "ספק" in c_raw:
            return c
    return None

AW_COL = _find_supplier_column(df.columns)
STOCK_COL = _find_column_by_candidates(df.columns, ["STOCK"]) or (df.columns[79] if len(df.columns) > 79 else df.columns[-1])

# ==========================================================
# דשבורד רכש חוצה-פרויקטים - טוען ומנתח את בלוק ה-DEMAND (ביקוש חודשי לכל
# רכיב) מכל הפרויקטים הזמינים בבת אחת, לא רק מהפרויקט הפעיל כרגע. ה"ביקוש"
# מטופל כאן כ"מה שנרכוש בפועל" (בהתאם להנחיה) - בלי נטרול מול מלאי/ETA.
# הפונקציות האלה עצמאיות לחלוטין מהפרויקט הפעיל (df/df_raw הגלובליים למעלה) -
# כל אחת טוענת את הקובץ שלה בנפרד, כדי שאפשר יהיה להריץ את כולן יחד.
# ==========================================================
def _date_run_forward(named_cols, start, max_len=200):
    """רצף רציף של עמודות-תאריך החל מ-start. אם start עצמו לא תאריך, מנסה
    להתקדם עד 5 עמודות קדימה למקרה של כותרת/רווח קטן לפני התחלת הבלוק."""
    idxs = []
    i = start
    while i < len(named_cols) and len(idxs) < max_len:
        if _is_date_col(named_cols[i]):
            idxs.append(i)
            i += 1
        else:
            if idxs:
                break
            i += 1
            if i - start > 5:
                break
    return idxs

def _load_project_procurement(project_id, url):
    """טוען קובץ פרויקט בודד ומחזיר DataFrame של שורות ביקוש-חודשי-לרכיב:
    Project, PN, Description, Supplier, YearMonth, Quantity, Unit_Price, Value."""
    xls = pd.ExcelFile(url)
    sheet = xls.sheet_names[0]
    p_df_raw = pd.read_excel(url, header=None, sheet_name=sheet)
    p_header_row = _find_matrix_header_row(p_df_raw)
    if p_header_row is None:
        return pd.DataFrame()
    p_df = pd.read_excel(url, header=p_header_row, sheet_name=sheet)
    p_df.columns = [str(c).strip() if pd.notnull(c) else c for c in p_df.columns]

    p_pn_col = _find_column_by_candidates(p_df.columns, ["PN_ID"]) or p_df.columns[1]
    p_desc_col = _find_column_by_candidates(p_df.columns, ["DESCRIPTION"]) or p_df.columns[4]
    p_price_col = _find_price_column(p_df.columns)
    p_supplier_col = _find_supplier_column(p_df.columns)

    demand_hit = _find_text_cell(p_df_raw, "DEMAND")
    demand_cols_idx = _date_run_forward(p_df.columns, demand_hit[1]) if demand_hit is not None else []

    if not demand_cols_idx:
        # נפילה בטוחה: אין תווית 'DEMAND' מפורשת - מחפשים את בלוק התאריכים
        # השני (אחרי בלוק ה-PLAN הראשון שמיד אחרי עץ ה-BOM), שבפועל מייצג
        # אותו סוג נתונים בהרבה מהקבצים.
        p_dlp = _find_desc_level_pn_cols(p_df_raw)
        if p_dlp:
            plan_start = (p_dlp["qty"] + 1) if p_dlp["qty"] is not None else (p_dlp["pn"] + 1)
            plan_run = _date_run_forward(p_df.columns, plan_start)
            if plan_run:
                after_plan = plan_run[-1] + 1
                for skip in range(0, 6):
                    demand_cols_idx = _date_run_forward(p_df.columns, after_plan + skip)
                    if demand_cols_idx:
                        break

    if not demand_cols_idx:
        return pd.DataFrame()
    demand_cols = [p_df.columns[i] for i in demand_cols_idx]

    # מחירים ידניים שמורים - נמשכים בנפרד לכל פרויקט (project_id), כדי שאותו
    # מק"ט יוכל להיות במחיר ידני שונה בכל פרויקט בלי להתערבב.
    try:
        p_inv_cache = fetch_all_inventory_records(project_id)
    except Exception:
        p_inv_cache = {}

    records = []
    for _, row in p_df.iterrows():
        pn = str(row.get(p_pn_col, "")).strip()
        if not pn or pn == "nan":
            continue
        desc = str(row.get(p_desc_col, ""))
        supplier = str(row.get(p_supplier_col, "")).strip() if p_supplier_col else ""
        if supplier.lower() in ["nan", "none", ""]:
            supplier = ""
        file_unit_price = safe_num(row.get(p_price_col)) if p_price_col else 0.0
        _p_rec = p_inv_cache.get(pn)
        unit_price = file_unit_price
        if _p_rec:
            _p_override = _p_rec.get("unit_price")
            if _p_override is not None and str(_p_override).strip() not in ["", "None", "nan"]:
                try:
                    unit_price = float(_p_override)
                except (TypeError, ValueError):
                    pass
        for c in demand_cols:
            qty = safe_num(row.get(c))
            if qty > 0:
                try:
                    ym = pd.to_datetime(c).strftime("%Y-%m")
                except Exception:
                    continue
                records.append({
                    "Project": project_id, "PN": pn, "Description": desc,
                    "Supplier": supplier or "לא צוין", "YearMonth": ym,
                    "Quantity": qty, "Unit_Price": unit_price, "Value": qty * unit_price
                })
    return pd.DataFrame(records)

@st.cache_data(ttl=600, show_spinner=False)
def load_all_projects_procurement(projects_dict):
    """טוען את כל הפרויקטים הזמינים (projects_dict: {שם: קישור}) ומאחד לטבלת
    רכש אחת. פרויקט שנכשל בטעינה מדולג עם אזהרה, לא מפיל את כל הדשבורד."""
    all_dfs = []
    failed = []
    for p_name, p_url in projects_dict.items():
        try:
            p_result = _load_project_procurement(p_name, p_url)
            if not p_result.empty:
                all_dfs.append(p_result)
            else:
                failed.append((p_name, "לא נמצאו נתוני ביקוש חודשי בקובץ"))
        except Exception as e:
            failed.append((p_name, str(e)))
    combined = pd.concat(all_dfs, ignore_index=True) if all_dfs else pd.DataFrame(
        columns=["Project", "PN", "Description", "Supplier", "YearMonth", "Quantity", "Unit_Price", "Value"]
    )
    return combined, failed

if ITEM_TYPE_COL is None:
    st.warning("⚠️ לא נמצאה עמודת 'סיווג פריט' (ITEM_TYPE) בקובץ - סינון/עריכת סוג פריט לא יהיו זמינים לפרויקט הזה.")
if PRICE_COL is None:
    st.warning("⚠️ לא נמצאה עמודת מחיר בקובץ - חישובי ערך כספי לא יהיו זמינים לפרויקט הזה.")
    PRICE_COL = df.columns[-1]
if AW_COL is None:
    AW_COL = df.columns[-1]

# ==========================================================
# פריסת עץ ה-BOM ישירות מהקובץ - דינמי לחלוטין: לא משנה כמה הרכבות יש,
# כמה רמות עומק יש בעץ המוצר, או איפה בדיוק בגיליון נמצא הבלוק. מאתרים
# את עמודות DESC/LEVEL/PN לפי הכותרות בפועל, וסורקים את כל גובה הגיליון
# (לא רק "אחרי" שורת הכותרת - בקבצים מסוימים המידע יושב לפניה) אחר שורות
# תקינות. אם אין עמודת כמות צמודה, הכמות-להורה נגזרת ממטריצת ה-where-used.
# ==========================================================
ASSEMBLY_BOM_TREE = {}
ASSEMBLY_CHILDREN = {}
# מפתח (הורה, ילד) -> הכמות הגולמית שנקראה ישירות מהקובץ עבור הקשר הספציפי הזה
# (מעמודת הכמות המפורשת, או מתא במטריצת ה-where-used). זה עדיין *לא* הכמות
# הסופית לחישובי MRP - היא רק שלב ביניים, כי אותה הרכבה יכולה להופיע תחת כמה
# הורים שונים (כמה ענפים בעץ), וצריך לסכום את כל המופעים שלה כדי לקבל את
# הכמות האמיתית ליחידת-מערכת (ראה בהמשך, אחרי שהעץ נבנה במלואו).
_RAW_EDGE_VALUE = {}
ordered_assemblies_from_excel = []

def _fill_edge_qty_from_matrix(children, df, pn_col, raw_edge_value):
    """ממלא raw_edge_value (כמות גולמית לקשר) לפי מטריצת ה-where-used, עבור כל
    קשר (הורה,ילד) בנפרד - כולל המקרה שבו אותו ילד מופיע תחת כמה הורים שונים.
    לא נוגע בקשרים שכבר יש להם ערך ממקור אחר (למשל עמודת כמות מפורשת)."""
    df_pnid = df[pn_col].astype(str).str.strip()
    match_cache = {}
    for parent_pn, kids in children.items():
        for child_pn in kids:
            if (parent_pn, child_pn) in raw_edge_value:
                continue
            if child_pn not in match_cache:
                match_cache[child_pn] = df_pnid[df_pnid == child_pn].index
            match_idx = match_cache[child_pn]
            q = 1.0
            if len(match_idx) > 0 and parent_pn in df.columns:
                q = safe_num(df.loc[match_idx[0], parent_pn], default=1.0)
                if q <= 0:
                    q = 1.0
            raw_edge_value[(parent_pn, child_pn)] = q

_dlp = _find_desc_level_pn_cols(df_raw)
if _dlp is not None:
    # --- אסטרטגיה 1: בלוק DESC/LEVEL/PN נפרד (מחפשים אותו בכל גובה הגיליון) ---
    try:
        _desc_c, _level_c, _pn_c, _qty_c = _dlp["desc"], _dlp["level"], _dlp["pn"], _dlp["qty"]
        _rows = []
        for _r in range(df_raw.shape[0]):
            _pn = df_raw.iloc[_r, _pn_c]
            _level = df_raw.iloc[_r, _level_c]
            if pd.isna(_pn) or str(_pn).strip() == "" or str(_level).strip().upper() == "LEVEL":
                continue
            try:
                _level = int(_level)
            except (ValueError, TypeError):
                continue
            _desc = df_raw.iloc[_r, _desc_c]
            _qty_raw = df_raw.iloc[_r, _qty_c] if _qty_c is not None else None
            _rows.append((_r, str(_pn).strip(), _level, _desc, _qty_raw))
        _rows.sort(key=lambda x: x[0])  # לפי סדר הופעה אמיתי בגיליון - שומר על סדר המעבר בעץ

        _level_stack = []
        for _r, _pn, _level, _desc, _qty_raw in _rows:
            while _level_stack and _level_stack[-1][0] >= _level:
                _level_stack.pop()
            _parent_pn = _level_stack[-1][1] if _level_stack else None

            if _qty_c is not None and _parent_pn is not None:
                _raw_val = safe_num(_qty_raw, default=1.0)
                # אם יש כבר ערך לקשר הזה (למשל אם אותו PN מופיע פעמיים תחת אותו
                # הורה בדיוק - נדיר, אבל ליתר ביטחון), מצרפים ולא דורסים.
                _RAW_EDGE_VALUE[(_parent_pn, _pn)] = _RAW_EDGE_VALUE.get((_parent_pn, _pn), 0.0) + _raw_val

            ASSEMBLY_BOM_TREE[_pn] = {
                "desc": str(_desc),
                "level": _level,
                "qty_per_system": None,   # יחושב בהמשך, אחרי שכל העץ (וכל הענפים) ידועים
                "qty_per_parent": None,   # יחושב בהמשך
                "parent": _parent_pn
            }
            if _pn not in ordered_assemblies_from_excel:
                ordered_assemblies_from_excel.append(_pn)
            if _parent_pn:
                ASSEMBLY_CHILDREN.setdefault(_parent_pn, []).append(_pn)
            _level_stack.append((_level, _pn))

        if _qty_c is None and ASSEMBLY_BOM_TREE:
            _fill_edge_qty_from_matrix(ASSEMBLY_CHILDREN, df, PN_COL, _RAW_EDGE_VALUE)
    except Exception as e:
        st.error(f"Error parsing BOM (אסטרטגיה 1): {e}")
        ASSEMBLY_BOM_TREE = {}
        ASSEMBLY_CHILDREN = {}
        _RAW_EDGE_VALUE = {}
        ordered_assemblies_from_excel = []

if not ASSEMBLY_BOM_TREE:
    # --- אסטרטגיה 2: אין בלוק DESC/LEVEL/PN נפרד - מנסים לשחזר את העץ ישירות
    # מתוך מטריצת ה-where-used, לפי שתי שורות-תיוג ('DESC' ו-'LEVEL'/'LEVAL')
    # שיושבות מעל שורת הכותרת הראשית ומתארות כל עמודת הרכבה. ---
    _strategy2_matrix_end_idx = None
    try:
        _mheader_row = _find_matrix_header_row(df_raw)
        _desc_row, _level_row = (None, None)
        if _mheader_row is not None:
            _desc_row, _level_row = _find_matrix_level_desc_rows(df_raw, _mheader_row)
        if _desc_row is not None and _level_row is not None:
            _matrix_cols = _find_matrix_col_run(df_raw, _level_row)
            _level_stack = []
            for _c in _matrix_cols:
                _level = int(df_raw.iat[_level_row, _c])
                # קוראים את המק"ט מהגיליון הגולמי (df_raw), לא מ-df.columns - כי אם
                # אותו מק"ט מופיע פעמיים כעמודת הרכבה (הרכבה משותפת בין שני ענפים/
                # שתי מערכות), פנדס מוסיף אוטומטית סיומת ל-df.columns (למשל
                # '...-001.1') שהייתה יוצרת בטעות שני פריטים נפרדים במקום אחד.
                _pn = str(df_raw.iat[_mheader_row, _c]).strip()
                _desc = df_raw.iat[_desc_row, _c]
                while _level_stack and _level_stack[-1][0] >= _level:
                    _level_stack.pop()
                _parent_pn = _level_stack[-1][1] if _level_stack else None
                ASSEMBLY_BOM_TREE[_pn] = {
                    "desc": str(_desc), "level": _level,
                    "qty_per_system": None, "qty_per_parent": None, "parent": _parent_pn
                }
                if _pn not in ordered_assemblies_from_excel:
                    ordered_assemblies_from_excel.append(_pn)
                if _parent_pn:
                    ASSEMBLY_CHILDREN.setdefault(_parent_pn, []).append(_pn)
                _level_stack.append((_level, _pn))
            if _matrix_cols:
                _strategy2_matrix_end_idx = _matrix_cols[-1] + 1
            if ASSEMBLY_BOM_TREE:
                _fill_edge_qty_from_matrix(ASSEMBLY_CHILDREN, df, PN_COL, _RAW_EDGE_VALUE)
    except Exception as e:
        st.error(f"Error parsing BOM (אסטרטגיה 2): {e}")
        ASSEMBLY_BOM_TREE = {}
        ASSEMBLY_CHILDREN = {}
        _RAW_EDGE_VALUE = {}
        ordered_assemblies_from_excel = []
else:
    _strategy2_matrix_end_idx = None

# אם שתי האסטרטגיות נכשלו, אין טעם להמשיך - כל שאר האפליקציה תלויה בעץ הזה
# ותיכשל בהמשך בצורה מבלבלת הרבה יותר מהודעה ברורה אחת כאן.
if not ASSEMBLY_BOM_TREE:
    st.error(
        "לא הצלחתי לפרוס את עץ ה-BOM עבור הפרויקט הזה - לא נמצא בלוק 'DESC/LEVEL/PN' נפרד, "
        "וגם לא זוהו שתי שורות-תיוג 'DESC'/'LEVEL' מעל שורת הכותרת הראשית של מטריצת ההרכבות. "
        "שלח את הקובץ הזה כדי שנבדוק מה שונה במבנה שלו."
    )
    st.stop()

# ==========================================================
# "מקדם מערכת" (כמה מהרכבה הזו נדרשים ליחידת-מערכת אחת) וכמות ליחידת-הורה,
# ספציפית לכל קשר. הכמות הגולמית שנקראה מהקובץ (_RAW_EDGE_VALUE) היא כבר
# ביחס ליחידת-מערכת ישירות (כך גם עמודת ה-QTY המפורשת בקבצים מהסוג הישן,
# וכך גם עמודת ה-SUM שקיימת בהרבה מקבצי המטריצה - שתיהן לא דורשות הכפלה
# נוספת בשרשרת ההורים). לכן מקדם המערכת של פריט הוא פשוט סכום כל הכמויות
# הגולמיות שהגיעו אליו מכל ההורים שלו (ענף אחד = סכום עם איבר בודד; שני
# ענפים עם יחס 2 כל אחד = סכום 4, לא 2*2). שורש (בלי הורה בכלל) = 1.0.
# כמות ליחידת-הורה (המשמשת בפועל בחישובי CTB/קיבולת מקסימלית/הקצאה משותפת)
# מחולקת מהכמות הגולמית של הקשר במקדם המערכת של ההורה עצמו - לא מוכפלת בו.
# ==========================================================
_PARENTS_OF = {}
for (_p, _c), _raw in _RAW_EDGE_VALUE.items():
    _PARENTS_OF.setdefault(_c, []).append((_p, _raw))

QTY_PER_SYSTEM = {}
for _pn in ASSEMBLY_BOM_TREE:
    _parents = _PARENTS_OF.get(_pn, [])
    QTY_PER_SYSTEM[_pn] = sum(_raw for _p, _raw in _parents) if _parents else 1.0
    ASSEMBLY_BOM_TREE[_pn]["qty_per_system"] = QTY_PER_SYSTEM[_pn]

ASSEMBLY_EDGE_QTY = {}
for (_p, _c), _raw in _RAW_EDGE_VALUE.items():
    _parent_factor = QTY_PER_SYSTEM.get(_p, 1.0) or 1.0
    ASSEMBLY_EDGE_QTY[(_p, _c)] = _raw / _parent_factor

for _pn, _node in ASSEMBLY_BOM_TREE.items():
    # qty_per_parent לתצוגה - הכמות מול ההורה שנשמר בפועל בעץ (הענף המייצג).
    # לחישובי MRP בפועל תמיד יש להשתמש ב-ASSEMBLY_EDGE_QTY לפי (הורה,ילד) ולא בזה.
    _rep_parent = _node["parent"]
    _node["qty_per_parent"] = ASSEMBLY_EDGE_QTY.get((_rep_parent, _pn), 1.0) if _rep_parent is not None else 1.0

ASSEMBLY_SYSTEM_FACTORS = {
    pn: factor for pn, factor in QTY_PER_SYSTEM.items() if abs(factor - 1.0) > 1e-9
}

# ==========================================================
# עמודות מטריצת ה"היכן משתמשים" (where-used) - נגזרות דינמית מתוך עץ ה-BOM
# שזוהה למעלה: כל עמודה בקובץ ששם הכותרת שלה מופיע בעץ, נחשבת עמודת הרכבה.
# לא תלוי במיקום קבוע (10:36) ולא במספר עמודות המטא-דאטה שקודמות למטריצה,
# ששונה בין קבצים (בקובץ הישן 10, בקובץ החדש 8).
# ==========================================================
if ASSEMBLY_BOM_TREE:
    ASSEMBLY_COLS = [c for c in df.columns if str(c).strip() in ASSEMBLY_BOM_TREE]
else:
    # נפילה בטוחה למבנה הידוע היסטורית, אם פריסת העץ נכשלה לחלוטין
    ASSEMBLY_COLS = df.columns[10:36].tolist() if len(df.columns) > 36 else []

valid_assemblies = list(ASSEMBLY_COLS)

# רמת ההרכבה נלקחת ישירות מהעץ הדינמי שזוהה למעלה (מקור אמת יחיד)
assembly_levels = {}
for col in valid_assemblies:
    assembly_levels[col] = ASSEMBLY_BOM_TREE.get(col, {}).get("level", 0)

valid_assemblies = sorted(
    valid_assemblies, 
    key=lambda x: ordered_assemblies_from_excel.index(x) if x in ordered_assemblies_from_excel else 999
)

# ==========================================================
# זיהוי דינמי של גבולות בלוקי החודשים (תוכנית עבודה / DEMAND / אספקה-ETA)
# לפי מיקום אמיתי בקובץ, ולא לפי טווח עמודות קבוע - כך שגם אופק תכנון שונה
# (פחות/יותר חודשים), תאריך התחלה שונה (כל שנה/חודש), ותוויות שונות
# ("PROJECT PLAN" מול "WORK PLAN") ימשיכו לעבוד נכון בלי לגעת בקוד.
# ==========================================================
_demand_hit = _find_text_cell(df_raw, "DEMAND")
_plan_start_idx = (_dlp["qty"] + 1) if (_dlp and _dlp["qty"] is not None) else ((_dlp["pn"] + 1) if _dlp else _strategy2_matrix_end_idx)
_demand_start_idx = _demand_hit[1] if _demand_hit else None

if _plan_start_idx is not None and _demand_start_idx is not None and _demand_start_idx > _plan_start_idx:
    MONTH_COLS = df.columns[_plan_start_idx:_demand_start_idx].tolist()
elif _plan_start_idx is not None:
    # אין תווית DEMAND - לוקחים את הרצף הרציף של עמודות-תאריך שמתחיל מיד אחרי PN/QTY
    _i, _found = _plan_start_idx, []
    while _i < len(df.columns) and _is_date_col(df.columns[_i]):
        _found.append(_i)
        _i += 1
    MONTH_COLS = df.columns[_found[0]:_found[-1] + 1].tolist() if _found else []
else:
    # נפילה בטוחה למבנה הידוע היסטורית
    MONTH_COLS = df.columns[108:132].tolist() if len(df.columns) > 132 else []

def _validate_month_headers(month_cols):
    warnings_list = []
    ym_list = []
    for c in month_cols:
        if pd.notnull(c):
            try:
                ym_list.append(pd.to_datetime(c).strftime("%Y-%m"))
            except Exception:
                pass
    if len(ym_list) != len(set(ym_list)):
        warnings_list.append(
            "⚠️ נמצאו חודשים כפולים בכותרת עמודות ה-MRP הראשיות. "
            "מבנה קובץ ה-Excel כנראה השתנה מאז התיקון האחרון - יש לבדוק ידנית "
            "לפני שסומכים על תוצאות המערכת."
        )
    if len(ym_list) == 0:
        warnings_list.append(
            "⚠️ לא זוהו עמודות חודשים תקינות בבלוק תוכנית העבודה. ייתכן שמבנה הקובץ השתנה."
        )
    return warnings_list

_month_header_warnings = _validate_month_headers(MONTH_COLS)
for _w in _month_header_warnings:
    st.warning(_w)

# ==========================================================
# תצוגת אימות בסיידבר - כדי שתראה מיד מה בפועל נטען, ותוכל לוודא שזה אכן
# הפרויקט שהתכוונת אליו (במיוחד חשוב אם דרסת קובץ קיים באותה כתובת GitHub
# בלי להוסיף אותו כפרויקט חדש - התוכן ישתנה, אבל ה-project_id יישאר הישן).
# ==========================================================
with st.sidebar:
    _top_level_descs = [info["desc"] for info in ASSEMBLY_BOM_TREE.values() if info["level"] == 0]
    st.caption(
        f"✅ נטענו {len(ASSEMBLY_BOM_TREE)} פריטים בעץ ({len(valid_assemblies)} הרכבות) | "
        f"{len(MONTH_COLS)} חודשים"
        + (f" ({str(MONTH_COLS[0])[:7]} עד {str(MONTH_COLS[-1])[:7]})" if MONTH_COLS else "")
    )
    if _top_level_descs:
        st.caption("מוצרים ברמה 0: " + ", ".join(_top_level_descs[:4]) + (" ..." if len(_top_level_descs) > 4 else ""))
    st.caption(f"מזהה פרויקט (project_id): `{PROJECT_ID}`")

if "custom_assembly_plan_df" not in st.session_state:
    cloud_plan = fetch_cloud_assembly_plan(PROJECT_ID)
    if not cloud_plan.empty:
        st.session_state["custom_assembly_plan_df"] = cloud_plan
    else:
        # אורך אופק התכנון וכל עמודות ה-PN/תאריך נלקחים מהזיהוי הדינמי למעלה,
        # ולא ננעלים לטווח שורות/עמודות קבוע - כך שזה עובד לכל פרויקט.
        plan_col_start = _plan_start_idx if _plan_start_idx is not None else 108
        pn_col_idx = _dlp["pn"] if _dlp else 106
        if len(MONTH_COLS) > 0:
            header_dates = list(MONTH_COLS)
        else:
            header_dates = df_raw.iloc[2, 108:132].values if df_raw.shape[1] > 132 else []
        plan_rows = []

        for r in range(0, df_raw.shape[0]):
            asm_pn = df_raw.iloc[r, pn_col_idx] if df_raw.shape[1] > pn_col_idx else None
            # מסננים לפי מק"טים שבאמת זוהו בעץ ה-BOM (לא כל ערך לא-ריק בעמודה הזו -
            # יש בקובץ שורות "רעש" כמו תוויות נוסחה שיושבות באותה עמודה בטעות)
            if pd.notnull(asm_pn) and str(asm_pn).strip() in ASSEMBLY_BOM_TREE:
                clean_asm_pn = str(asm_pn).strip()
                system_multiplier = ASSEMBLY_SYSTEM_FACTORS.get(clean_asm_pn, 1)

                for c_idx, date_val in enumerate(header_dates):
                    if pd.notnull(date_val):
                        qty = df_raw.iloc[r, plan_col_start + c_idx]
                        if pd.notnull(qty) and qty != '' and qty != 'NaN':
                            try:
                                q_val = float(qty)
                                if q_val > 0:
                                    dt = pd.to_datetime(date_val)
                                    ym_str = dt.strftime("%Y-%m")
                                    # כל החודשים בבלוק תוכנית העבודה שזוהה נכללים - בלי סינון
                                    # לפי תאריך-סף קבוע (שהיה נעול ל-2026-09 ולא רלוונטי לפרויקטים אחרים).
                                    displayed_build_qty = q_val * system_multiplier
                                    plan_rows.append({
                                        "Assembly_PN": clean_asm_pn,
                                        "YearMonth": ym_str,
                                        "Build_Qty": displayed_build_qty,
                                        "Raw_Build_Qty": q_val
                                    })
                            except:
                                pass
        st.session_state["custom_assembly_plan_df"] = pd.DataFrame(plan_rows)

assembly_plan_df = st.session_state["custom_assembly_plan_df"]

# ==========================================================
# בלוק "לוח האספקה" (ETA לפי חודש) - עד עכשיו הקוד היה מייצר תאריכים באופן מלאכותי
# (תמיד החל מינואר 2026), במקום לקרוא את התאריכים האמיתיים שכתובים בקובץ.
# זה שובר כל פרויקט שמתחיל בשנה/חודש אחרים. עכשיו קוראים את התאריכים בפועל
# מכותרת העמודות בקובץ (מיד אחרי עמודת ה-STOCK ועד עמודת ה-DESC), כך שזה
# יעבוד נכון לכל פרויקט, לא משנה מתי הוא מתחיל.
# ==========================================================
# ==========================================================
# בלוק "לוח האספקה" (ETA לפי חודש) - עד עכשיו הקוד היה מייצר תאריכים באופן מלאכותי
# (תמיד החל מינואר 2026, לפי מיקום עמודה בלבד), במקום לקרוא את התאריכים האמיתיים
# שכתובים בקובץ. זה שובר כל פרויקט שמתחיל בשנה/חודש אחרים. עכשיו מאתרים את הרצף
# הרציף הארוך ביותר של עמודות-תאריך שמסתיים ממש לפני בלוק ה-DESC/LEVEL/PN (בלי
# להניח את סדר העמודות היחסי בין "STOCK" לבלוק התאריכים - זה משתנה בין קבצים),
# וקוראים משם את התאריכים בפועל.
# ==========================================================
_supply_range = _longest_date_run_ending_before(df.columns, _dlp["desc"]) if _dlp else None
if _supply_range:
    _supply_block_start_idx, _supply_block_end_idx = _supply_range
else:
    # נפילה בטוחה למבנה הידוע היסטורית, אם לא אותר בלוק תאריכים
    _supply_block_start_idx, _supply_block_end_idx = 80, 104

def _build_supply_date_map():
    date_map = {}
    for col_pos in range(_supply_block_start_idx, _supply_block_end_idx):
        if col_pos >= len(df.columns):
            break
        try:
            dt = pd.to_datetime(df.columns[col_pos])
            if pd.notnull(dt):
                date_map[col_pos] = dt
        except Exception:
            pass
    plan_col_start = _plan_start_idx if _plan_start_idx is not None else 108
    for i, col in enumerate(MONTH_COLS):
        col_pos = plan_col_start + i
        if pd.notnull(col):
            try:
                date_map[col_pos] = pd.to_datetime(col)
            except Exception:
                pass
    return date_map

SUPPLY_DATE_MAP = _build_supply_date_map()


def get_base_mrp_eta_and_qty(pn):
    matching_rows = df_raw[df_raw.iloc[:, 1].astype(str).str.strip() == str(pn).strip()]
    if matching_rows.empty:
        return "בדיקה נדרשת", 0.0

    row_idx = matching_rows.index[0]
    for col_pos in sorted(SUPPLY_DATE_MAP.keys()):
        try:
            val = df_raw.iloc[row_idx, col_pos]
            q = safe_num(val)
            if q > 0:
                dt = SUPPLY_DATE_MAP[col_pos]
                return dt.strftime("%Y-%m"), q
        except Exception:
            pass
    return "בדיקה נדרשת", 0.0

def get_base_mrp_eta(pn):
    eta, _ = get_base_mrp_eta_and_qty(pn)
    return eta

def get_base_mrp_qty(pn):
    _, qty = get_base_mrp_eta_and_qty(pn)
    return qty

def get_cumulative_incoming_supply(pn, target_ym):
    matching_rows = df_raw[df_raw.iloc[:, 1].astype(str).str.strip() == str(pn).strip()]
    if matching_rows.empty:
        return 0.0
    row_idx = matching_rows.index[0]
    total = 0.0
    for col_pos, dt in SUPPLY_DATE_MAP.items():
        if col_pos >= _supply_block_end_idx:
            continue
        try:
            ym = dt.strftime("%Y-%m")
            if ym < target_ym:
                q = safe_num(df_raw.iloc[row_idx, col_pos])
                if q > 0:
                    total += q
        except Exception:
            pass
    return total

def get_component_available_by_month(pn, target_ym, inv_cache=None, wip_cache=None):
    if inv_cache is None:
        inv_cache = fetch_all_inventory_records(PROJECT_ID)
    if wip_cache is None:
        wip_cache = fetch_wip_records(PROJECT_ID)

    match = df[df[PN_COL].astype(str).str.strip() == pn]
    base_stock = safe_num(match.iloc[0][STOCK_COL]) if not match.empty else 0.0

    saved_add, manual_eta, _, _, _, _, _ = get_inventory_record(pn, inv_cache)
    manual_eta_ym = None
    if manual_eta and str(manual_eta).strip() not in ["", "None", "NaT", "nan"]:
        try:
            manual_eta_ym = pd.to_datetime(manual_eta).strftime("%Y-%m")
        except Exception:
            manual_eta_ym = None
    manual_stock_effective = saved_add if (manual_eta_ym is None or manual_eta_ym < target_ym) else 0.0

    incoming_supply = get_cumulative_incoming_supply(pn, target_ym)

    wip_committed = 0.0
    if not match.empty and wip_cache:
        row0 = match.iloc[0]
        for asm_col in ASSEMBLY_COLS:
            wip_qty = wip_cache.get(asm_col, 0.0)
            if wip_qty > 0 and asm_col in df.columns:
                qty_per = safe_num(row0.get(asm_col, 0.0))
                if qty_per > 0:
                    sys_factor = ASSEMBLY_SYSTEM_FACTORS.get(asm_col, 1)
                    wip_committed += (wip_qty / sys_factor) * qty_per

    return max(0.0, base_stock + manual_stock_effective + incoming_supply - wip_committed)

def get_first_supply_eta(pn, inv_cache=None):
    _, manual_eta, _, _, _, _, _ = get_inventory_record(pn, inv_cache)
    if manual_eta and str(manual_eta).strip() not in ["", "None", "NaT", "nan"]:
        return manual_eta
    return get_base_mrp_eta(pn)

# ==========================================================
# SIDEBAR FILTERS & WHAT-IF CONTROLS
# ==========================================================
st.sidebar.header("⚙️ הגדרות מערכת וחיבור")
webhook_url = st.sidebar.text_input("🔗 Teams / Slack Webhook URL (אופציונלי)", value="")
supplier_options = ["אופק", "ספק פנימי", "רכש אחר", "אחר"]

st.sidebar.header("🔍 מסננים מתקדמים")

if st.sidebar.button("🧹 איפוס כל המסננים (Clear All)"):
    keys_to_clear = ["selected_month_label", "num_months_ahead", "selected_level", "selected_assembly", "selected_item_type", "selected_search_item", "selected_aw"]
    for k in keys_to_clear:
        if k in st.session_state:
            del st.session_state[k]
    st.rerun()

month_options = {}
for m in MONTH_COLS:
    if pd.notnull(m):
        try:
            dt = pd.to_datetime(m)
            m_ym = dt.strftime("%Y-%m")
            if m_ym >= "2026-09":
                month_options[dt.strftime("%B %Y (שנה-חודש: %Y-%m)")] = m
        except:
            pass

if not month_options:
    for m in MONTH_COLS:
        if pd.notnull(m):
            try:
                dt = pd.to_datetime(m)
                if dt.month >= 9 or dt.strftime("%Y-%m") >= "2026-09":
                    month_options[dt.strftime("%B %Y (שנה-חודש: %Y-%m)")] = m
            except:
                pass

selected_month_label = st.sidebar.selectbox("בחר חודש לניתוח חוסרים", list(month_options.keys()), key="selected_month_label")
selected_month_col = month_options[selected_month_label]

try:
    selected_ym = pd.to_datetime(selected_month_col).strftime("%Y-%m")
except:
    selected_ym = str(selected_month_col)[:7]

num_months_ahead = st.sidebar.slider("📅 טווח מבט קדימה במספר חודשים", min_value=1, max_value=6, value=1, step=1, format="%d חודשים קדימה", key="num_months_ahead")

level_options = ["הכל"] + sorted(list(set(str(assembly_levels[c]) for c in valid_assemblies)), key=lambda x: int(x) if x.isdigit() else 0)
selected_level = st.sidebar.selectbox("סינון לפי רמת עץ (BOM Level)", level_options, key="selected_level")

assembly_mapping = {"הכל": "הכל"}
filtered_assembly_cols = []
for col in valid_assemblies:
    try:
        lvl = str(assembly_levels.get(col, 0))
        desc = ASSEMBLY_BOM_TREE.get(col, {}).get("desc", "")
        if selected_level == "הכל" or lvl == selected_level:
            filtered_assembly_cols.append(col)
            assembly_mapping[col] = f"[רמה {lvl}] {str(col)} - {str(desc)}"
    except:
        filtered_assembly_cols.append(col)
        assembly_mapping[col] = col

selected_assembly = st.sidebar.selectbox(
    "בחר הרכבה ספציפית לדשבורד",
    ["הכל"] + filtered_assembly_cols,
    format_func=lambda x: assembly_mapping.get(x, x),
    key="selected_assembly"
)

item_types = df[ITEM_TYPE_COL].dropna().unique().tolist() if ITEM_TYPE_COL in df.columns else []
_override_item_types = set(
    str(v.get("item_type", "")).strip()
    for v in fetch_all_inventory_records(PROJECT_ID).values()
    if str(v.get("item_type", "")).strip()
)
item_types = sorted(set(item_types) | _override_item_types)
selected_item_type = st.sidebar.selectbox("בחר סוג פריט (עמודה AS)", ["הכל"] + item_types, key="selected_item_type")

aw_values = df[AW_COL].dropna().astype(str).unique().tolist() if AW_COL in df.columns else []
aw_values = sorted(list(set(aw_values)))
selected_aw = st.sidebar.selectbox("סינון לפי ספק או BFE ", ["הכל"] + aw_values, key="selected_aw")

item_choices = ["הכל"] + sorted([f"{str(r[PN_COL]).strip()} - {str(r[DESC_COL])}" for _, r in df.iterrows() if pd.notnull(r[PN_COL])])
selected_search_item = st.sidebar.selectbox("🔎 חיפוש מהיר (בחר או הקלד מק'ט/תיאור)", item_choices, key="selected_search_item")
search_pn = selected_search_item.split(" - ")[0] if selected_search_item != "הכל" else "הכל"

# ==========================================================
# FILE UPLOADS & TEMPLATES
# ==========================================================
st.sidebar.divider()
st.sidebar.markdown("##### 📥 עדכון ETA וכמות אספקה מקובץ ספק")

eta_template_df = pd.DataFrame(columns=["PN", "ETA", "Qty"])
eta_template_output = io.BytesIO()
with pd.ExcelWriter(eta_template_output, engine='openpyxl') as writer:
    eta_template_df.to_excel(writer, index=False, sheet_name='ETA_Template')
st.sidebar.download_button(
    label="📄 הורד תבנית Excel לעדכון ETA",
    data=eta_template_output.getvalue(),
    file_name="ETA_Update_Template.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)

uploaded_eta_file = st.sidebar.file_uploader("העלה קובץ ETA (עמודות: PN, ETA, Qty)", type=["xlsx", "xls"], key="eta_uploader")
if uploaded_eta_file is not None:
    try:
        eta_df_sup = pd.read_excel(uploaded_eta_file)
        if st.sidebar.button("⚡ עדכן ETA וכמות אספקה"):
            eta_count = 0
            for _, s_row in eta_df_sup.iterrows():
                p_code = str(s_row.iloc[0]).strip()
                new_eta = str(s_row.iloc[1]).strip() if len(s_row) > 1 and pd.notnull(s_row.iloc[1]) else ""
                new_supply_qty = float(s_row.iloc[2]) if len(s_row) > 2 and pd.notnull(s_row.iloc[2]) else 0.0

                if p_code and p_code != 'nan' and new_eta and new_eta not in ["nan", "NaT", "None"]:
                    curr_stock, _, curr_status, curr_sup, curr_comm, _, _ = get_inventory_record(p_code)
                    if not curr_sup:
                        curr_sup = get_effective_supplier(p_code)
                    updated_total_stock = curr_stock + new_supply_qty if new_supply_qty > 0 else curr_stock

                    save_inventory_record(
                        pn=p_code,
                        added_stock=updated_total_stock,
                        eta=new_eta,
                        status=curr_status if curr_status != "פתוח" else "הוזמן",
                        supplier=curr_sup,
                        comment=f"{curr_comm} | אספקה בסך {new_supply_qty} בתאריך ETA {new_eta} מקובץ ספק",
                        updated_by="ETA & Qty File Upload",
                        webhook_url=webhook_url
                    )
                    eta_count += 1
            st.sidebar.success(f"עודכנו בהצלחה ETA וכמויות אספקה עבור {eta_count} שורות!")
    except Exception as e:
        st.sidebar.error(f"שגיאה בקריאת קובץ ה-ETA: {e}")

# ==========================================================
# OPTIMIZED SHORTAGE CALCULATION
# ==========================================================
all_ym_list = sorted(list(set(assembly_plan_df["YearMonth"].unique())))
start_idx = 0
for idx, ym in enumerate(all_ym_list):
    if ym >= selected_ym:
        start_idx = idx
        break
selected_target_yms = all_ym_list[start_idx:start_idx + num_months_ahead]
if not selected_target_yms:
    selected_target_yms = [selected_ym]

@st.cache_data(ttl=60)
def calculate_mrp_breakdown_cached(target_yms_tuple, sim_extra_stock_items_tuple, active_plan_df):
    sim_extra_stock_dict = dict(sim_extra_stock_items_tuple)
    inv_cache = fetch_all_inventory_records(PROJECT_ID)
    wip_cache = fetch_wip_records(PROJECT_ID)

    target_month_cols_map = {}
    for m_c in MONTH_COLS:
        if pd.notnull(m_c):
            try:
                m_dt_ym = pd.to_datetime(m_c).strftime("%Y-%m")
                if m_dt_ym in target_yms_tuple:
                    target_month_cols_map[m_dt_ym] = m_c
            except:
                pass

    temp_df = df.copy()
    shortage_records = {}

    for idx, row in temp_df.iterrows():
        pn = str(row[PN_COL]).strip()
        saved_stock_add, manual_eta, _, _, _, _, _ = get_inventory_record(pn, inv_cache)
        sim_val = sim_extra_stock_dict.get(pn, 0.0)

        manual_eta_ym = None
        if manual_eta and str(manual_eta).strip() not in ["", "None", "NaT", "nan"]:
            try:
                manual_eta_ym = pd.to_datetime(manual_eta).strftime("%Y-%m")
            except Exception:
                manual_eta_ym = None

        max_shortage_val = 0.0
        is_short_or = False

        for ym in target_yms_tuple:
            col_name = target_month_cols_map.get(ym)
            if col_name and col_name in temp_df.columns:
                mrp_val = safe_num(row[col_name])

                stock_arrived_by_this_month = (manual_eta_ym is None) or (manual_eta_ym < ym)
                effective_addition = (saved_stock_add if stock_arrived_by_this_month else 0.0) + sim_val

                effective_mrp_val = mrp_val + effective_addition if mrp_val < 0 else mrp_val

                if effective_mrp_val < 0:
                    is_short_or = True
                    sh_qty = abs(effective_mrp_val)
                    if sh_qty > max_shortage_val:
                        max_shortage_val = sh_qty

        if is_short_or:
            shortage_records[idx] = max_shortage_val

    temp_df['Monthly_Balance'] = temp_df.index.map(lambda i: -shortage_records[i] if i in shortage_records else 1.0)

    mrp_shortages = temp_df[temp_df['Monthly_Balance'] < 0].copy()
    mrp_shortages['Total_MRP_Shortage'] = mrp_shortages['Monthly_Balance'].abs()

    month_plan = active_plan_df[active_plan_df["YearMonth"].isin(target_yms_tuple)]
    plan_dict = month_plan.groupby("Assembly_PN")["Raw_Build_Qty"].sum().to_dict()

    for asm_wip, wip_qty in wip_cache.items():
        if wip_qty > 0 and asm_wip in plan_dict:
            sys_factor = ASSEMBLY_SYSTEM_FACTORS.get(asm_wip, 1)
            raw_wip_qty = wip_qty / sys_factor
            plan_dict[asm_wip] = max(0.0, plan_dict[asm_wip] - raw_wip_qty)

    breakdown_rows = []
    reference_ym = max(target_yms_tuple) if target_yms_tuple else None

    for idx, row in mrp_shortages.iterrows():
        pn = str(row[PN_COL]).strip()
        desc = str(row[DESC_COL])
        original_item_type = str(row[ITEM_TYPE_COL]) if ITEM_TYPE_COL in temp_df.columns else ""
        item_type = get_effective_item_type(pn, original_item_type, inv_cache)

        if reference_ym:
            stock = get_component_available_by_month(pn, reference_ym, inv_cache, wip_cache) + sim_extra_stock_dict.get(pn, 0.0)
        else:
            base_stock = safe_num(row[STOCK_COL])
            saved_stock_add, _, _, _, _, _, _ = get_inventory_record(pn, inv_cache)
            stock = base_stock + saved_stock_add + sim_extra_stock_dict.get(pn, 0.0)

        total_mrp_shortage = row['Total_MRP_Shortage']
        _, _, item_status, current_sup, _, _, _ = get_inventory_record(pn, inv_cache)
        # אם לא נשמר שינוי ידני לספק, לוקחים את הספק בפועל מהקובץ (עמודת AW_COL) -
        # לא ברירת מחדל קבועה שלא בהכרח נכונה לפריט הזה.
        if not current_sup:
            current_sup = str(row.get(AW_COL, "")).strip() or "לא צוין"
        unit_price = get_effective_price(pn, safe_num(row[PRICE_COL]), inv_cache)
        shortage_value = total_mrp_shortage * unit_price

        mouser_link = f"https://www.mouser.co.il/c/?q={pn}"
        digikey_link = f"https://www.digikey.com/en/products/result?keywords={pn}"
        findchips_link = f"https://www.findchips.com/search/{pn}"

        added_for_this_pn = False
        for asm in ASSEMBLY_COLS:
            qty_per_asm = safe_num(row[asm])
            if qty_per_asm > 0:
                added_for_this_pn = True
                asm_raw_build = plan_dict.get(asm, 0.0)
                required_demand = qty_per_asm * asm_raw_build
                asm_desc = assembly_mapping.get(asm, asm)

                breakdown_rows.append({
                    "PN": pn, "Description": desc, "Item_Type": item_type, "Supplier": current_sup,
                    "Status": item_status, "Assembly": asm, "Assembly_Desc": asm_desc, "Qty_Per_Assembly": qty_per_asm,
                    "Assembly_Monthly_Build": asm_raw_build * ASSEMBLY_SYSTEM_FACTORS.get(asm, 1),
                    "Required_Demand": required_demand,
                    "Stock": stock, "Total_MRP_Shortage": total_mrp_shortage,
                    "Unit_Price": unit_price, "Shortage_Value": shortage_value,
                    "AW_Data": str(row.get(AW_COL, "")).strip(),
                    "חיפוש במאוזר": mouser_link, "חיפוש בדיגיקי": digikey_link, "חיפוש ב-Findchips": findchips_link
                })

        if not added_for_this_pn:
            breakdown_rows.append({
                "PN": pn, "Description": desc, "Item_Type": item_type, "Supplier": current_sup,
                "Status": item_status, "Assembly": "ללא שיוך", "Assembly_Desc": "ללא שיוך להרכבה", "Qty_Per_Assembly": 0,
                "Assembly_Monthly_Build": 0, "Required_Demand": 0, "Stock": stock, "Total_MRP_Shortage": total_mrp_shortage,
                "Unit_Price": unit_price, "Shortage_Value": shortage_value,
                "AW_Data": str(row.get(AW_COL, "")).strip(),
                "חיפוש במאוזר": mouser_link, "חיפוש בדיגיקי": digikey_link, "חיפוש ב-Findchips": findchips_link
            })

    res_df = pd.DataFrame(breakdown_rows)
    return res_df

def calculate_mrp_breakdown(sim_extra_stock=None, target_yms=None, plan_df_override=None):
    if sim_extra_stock is None:
        sim_extra_stock = {}
    if target_yms is None:
        target_yms = selected_target_yms
    active_plan = plan_df_override if plan_df_override is not None else assembly_plan_df
     
    res = calculate_mrp_breakdown_cached(tuple(target_yms), tuple(sorted(sim_extra_stock.items())), active_plan)
    res_df = res.copy()

    if not res_df.empty:
        if selected_item_type != "הכל":
            res_df = res_df[res_df["Item_Type"] == selected_item_type]
        if selected_assembly != "הכל":
            res_df = res_df[res_df["Assembly"] == selected_assembly]
        if search_pn != "הכל":
            res_df = res_df[res_df["PN"] == search_pn]
        if selected_aw != "הכל":
            res_df = res_df[res_df["AW_Data"] == selected_aw]

    return res_df

breakdown_df = calculate_mrp_breakdown(target_yms=selected_target_yms)

def compute_shared_executable_plan(target_yms, assemblies, inv_cache=None, wip_cache=None):
    if inv_cache is None:
        inv_cache = fetch_all_inventory_records(PROJECT_ID)
    if wip_cache is None:
        wip_cache = fetch_wip_records(PROJECT_ID)

    priority_order = sorted(assemblies, key=lambda a: (assembly_levels.get(a, 0), str(a)))
    chronological_yms = sorted(target_yms)
    consumed_so_far = {}
    result = {}

    for ym in chronological_yms:
        pool_cache = {}

        def get_pool(pn):
            if pn not in pool_cache:
                if pn in ASSEMBLY_BOM_TREE:
                    sub_wip = wip_cache.get(pn, 0.0)
                    sub_max_new, _ = compute_max_buildable(pn, ym, inv_cache, wip_cache)
                    total_available_by_month = sub_wip + sub_max_new
                else:
                    total_available_by_month = get_component_available_by_month(pn, ym, inv_cache, wip_cache)
                pool_cache[pn] = max(0.0, total_available_by_month - consumed_so_far.get(pn, 0.0))
            return pool_cache[pn]

        month_breakdown = calculate_mrp_breakdown(target_yms=[ym])

        for asm_col in priority_order:
            sys_factor = ASSEMBLY_SYSTEM_FACTORS.get(asm_col, 1)
            sub_plan_df = assembly_plan_df[(assembly_plan_df["YearMonth"] == ym) & (assembly_plan_df["Assembly_PN"] == asm_col)]
            
            raw_build = sub_plan_df["Raw_Build_Qty"].sum() if not sub_plan_df.empty else 0.0
            current_wip_qty = wip_cache.get(asm_col, 0.0)

            discrete_plan_build = raw_build * sys_factor
            remaining_plan_target = max(0.0, discrete_plan_build - current_wip_qty)

            if discrete_plan_build <= 0 and current_wip_qty <= 0:
                continue

            asm_shortages = month_breakdown[month_breakdown["Assembly"] == asm_col] if not month_breakdown.empty else pd.DataFrame()

            max_possible_new_build = remaining_plan_target
            limiting_components = []

            if not asm_shortages.empty and remaining_plan_target > 0:
                for _, s_row in asm_shortages.iterrows():
                    req_per = s_row["Qty_Per_Assembly"]
                    if req_per > 0:
                        comp_pn = str(s_row["PN"]).strip()
                        avail = get_pool(comp_pn)
                        possible = (avail / req_per) * sys_factor
                        if possible < max_possible_new_build - 1e-9:
                            max_possible_new_build = possible
                            limiting_components = [comp_pn]
                        elif abs(possible - max_possible_new_build) < 1e-9:
                            limiting_components.append(comp_pn)

            if remaining_plan_target > 0:
                for child_pn in ASSEMBLY_CHILDREN.get(asm_col, []):
                    # כמות ספציפית לקשר (asm_col -> child_pn), לא הכמות ה"גלובלית" של
                    # הפריט - חיוני כשאותו רכיב משותף בין כמה הרכבות-אב שונות.
                    qty_per_parent = ASSEMBLY_EDGE_QTY.get((asm_col, child_pn), 1.0)
                    if qty_per_parent <= 0:
                        continue
                    avail = get_pool(child_pn)
                    possible = avail / qty_per_parent
                    if possible < max_possible_new_build - 1e-9:
                        max_possible_new_build = possible
                        limiting_components = [child_pn]
                    elif abs(possible - max_possible_new_build) < 1e-9:
                        limiting_components.append(child_pn)

            net_executable_qty = max(0.0, min(remaining_plan_target, max_possible_new_build))

            if net_executable_qty > 0:
                if not asm_shortages.empty:
                    for _, s_row in asm_shortages.iterrows():
                        req_per = s_row["Qty_Per_Assembly"]
                        if req_per > 0:
                            comp_pn = str(s_row["PN"]).strip()
                            consumed_amount = (net_executable_qty / sys_factor) * req_per
                            pool_cache[comp_pn] = max(0.0, get_pool(comp_pn) - consumed_amount)
                            consumed_so_far[comp_pn] = consumed_so_far.get(comp_pn, 0.0) + consumed_amount
                for child_pn in ASSEMBLY_CHILDREN.get(asm_col, []):
                    qty_per_parent = ASSEMBLY_EDGE_QTY.get((asm_col, child_pn), 1.0)
                    if qty_per_parent <= 0:
                        continue
                    consumed_amount = qty_per_parent * net_executable_qty
                    pool_cache[child_pn] = max(0.0, get_pool(child_pn) - consumed_amount)
                    consumed_so_far[child_pn] = consumed_so_far.get(child_pn, 0.0) + consumed_amount

            result[(asm_col, ym)] = {
                "raw_build": discrete_plan_build,
                "gross_executable": net_executable_qty,
                "net_executable": net_executable_qty,
                "wip": current_wip_qty,
                "limiting_components": limiting_components,
            }

    return result

def check_hierarchical_ctb(asm_pn, requested_qty, target_ym, inv_cache=None, wip_cache=None, _visited=None, _claimed=None):
    if inv_cache is None:
        inv_cache = fetch_all_inventory_records(PROJECT_ID)
    if wip_cache is None:
        wip_cache = fetch_wip_records(PROJECT_ID)
    if _visited is None:
        _visited = set()
    # _claimed עוקב כמה מכל רכיב כבר "נתפס" על ידי ענפים אחרים באותה בדיקה -
    # חיוני כשרכיב משותף בין כמה ענפים (אותו PN תחת כמה הורים): בלי זה, כל ענף
    # היה בודק מול כל המלאי בנפרד ומפספס שהענף האחר כבר תובע חלק ממנו, מה
    # שמדווח על חוסר קטן מדי כשיש מלאי חלקי (לא אפס, לא מספיק) לרכיב המשותף.
    if _claimed is None:
        _claimed = {}
    if asm_pn in _visited or requested_qty <= 0:
        return []
    _visited.add(asm_pn)

    blockers = []

    if asm_pn in df.columns:
        sys_factor = ASSEMBLY_SYSTEM_FACTORS.get(asm_pn, 1)
        for _, row in df.iterrows():
            qty_per = safe_num(row[asm_pn])
            if qty_per <= 0:
                continue
            comp_pn = str(row[PN_COL]).strip()
            base_stock_check = safe_num(row[STOCK_COL])
            if base_stock_check >= 9000000:
                continue
            required = (requested_qty / sys_factor) * qty_per
            available = get_component_available_by_month(comp_pn, target_ym, inv_cache, wip_cache)
            already_claimed = _claimed.get(comp_pn, 0.0)
            remaining_available = max(0.0, available - already_claimed)
            _claimed[comp_pn] = already_claimed + required
            if remaining_available < required:
                blockers.append({
                    "assembly": asm_pn,
                    "assembly_desc": assembly_mapping.get(asm_pn, asm_pn),
                    "component": comp_pn,
                    "component_desc": str(row[DESC_COL]),
                    "required": required, "available": remaining_available,
                    "shortage": required - remaining_available
                })

    for child_pn in ASSEMBLY_CHILDREN.get(asm_pn, []):
        qty_per_parent = ASSEMBLY_EDGE_QTY.get((asm_pn, child_pn), 1.0)
        child_needed = requested_qty * qty_per_parent
        child_wip = wip_cache.get(child_pn, 0.0)
        net_needed = max(0.0, child_needed - child_wip)
        if net_needed > 0:
            blockers.extend(check_hierarchical_ctb(child_pn, net_needed, target_ym, inv_cache, wip_cache, _visited, _claimed))

    return blockers

def _max_buildable_from_direct_components(asm_pn, target_ym, inv_cache, wip_cache):
    max_qty = float('inf')
    limiting_component = None

    if asm_pn not in df.columns:
        return max_qty, limiting_component

    sys_factor = ASSEMBLY_SYSTEM_FACTORS.get(asm_pn, 1)

    for _, row in df.iterrows():
        qty_per = safe_num(row[asm_pn])
        if qty_per <= 0:
            continue
        comp_pn = str(row[PN_COL]).strip()
        base_stock_check = safe_num(row[STOCK_COL])
        if base_stock_check >= 9000000:
            continue
        avail = get_component_available_by_month(comp_pn, target_ym, inv_cache, wip_cache)
        possible = (avail / qty_per) * sys_factor
        if possible < max_qty:
            max_qty = possible
            limiting_component = comp_pn

    return max_qty, limiting_component

def _max_buildable_from_subassemblies(asm_pn, target_ym, inv_cache, wip_cache, visited):
    max_qty = float('inf')
    limiting_component = None

    for child_pn in ASSEMBLY_CHILDREN.get(asm_pn, []):
        qty_per_parent = ASSEMBLY_EDGE_QTY.get((asm_pn, child_pn), 1.0)
        if qty_per_parent <= 0:
            continue

        child_wip = wip_cache.get(child_pn, 0.0)
        child_max_new, _ = compute_max_buildable(child_pn, target_ym, inv_cache, wip_cache, visited)
        child_total_available = child_wip + child_max_new

        possible_from_child = child_total_available / qty_per_parent
        if possible_from_child < max_qty:
            max_qty = possible_from_child
            limiting_component = f"{child_pn} (תת-הרכבה)"

    return max_qty, limiting_component

def compute_max_buildable(asm_pn, target_ym, inv_cache=None, wip_cache=None, _visited=None):
    if inv_cache is None:
        inv_cache = fetch_all_inventory_records(PROJECT_ID)
    if wip_cache is None:
        wip_cache = fetch_wip_records(PROJECT_ID)
    if _visited is None:
        _visited = set()
    if asm_pn in _visited:
        return float('inf'), None
    _visited = _visited | {asm_pn}

    direct_qty, direct_limiter = _max_buildable_from_direct_components(asm_pn, target_ym, inv_cache, wip_cache)
    sub_qty, sub_limiter = _max_buildable_from_subassemblies(asm_pn, target_ym, inv_cache, wip_cache, _visited)

    if direct_qty <= sub_qty:
        max_qty, limiting_component = direct_qty, direct_limiter
    else:
        max_qty, limiting_component = sub_qty, sub_limiter

    if max_qty == float('inf'):
        max_qty = 0.0
    return max(0.0, max_qty), limiting_component

_exec_wip_cache = fetch_wip_records(PROJECT_ID)
_exec_unique_shortage_items = breakdown_df["PN"].nunique() if not breakdown_df.empty else 0
_exec_blocked_assemblies = breakdown_df[breakdown_df["Assembly"] != "ללא שיוך"]["Assembly"].nunique() if not breakdown_df.empty else 0
_exec_total_value = breakdown_df.drop_duplicates(subset=["PN"])["Shortage_Value"].sum() if (not breakdown_df.empty and "Shortage_Value" in breakdown_df.columns) else 0.0
_exec_active_wip = len([q for q in _exec_wip_cache.values() if q > 0])

st.markdown(f"""
<div class="exec-summary-strip">
    <div class="exec-stat">
        <div class="exec-stat-icon">💵</div>
        <div class="exec-stat-value">${_exec_total_value:,.0f}</div>
        <div class="exec-stat-label">ערך חוסרים כולל בטווח הנבחר</div>
    </div>
    <div class="exec-stat">
        <div class="exec-stat-icon">📦</div>
        <div class="exec-stat-value">{_exec_unique_shortage_items:,}</div>
        <div class="exec-stat-label">פריטים בחוסר</div>
    </div>
    <div class="exec-stat">
        <div class="exec-stat-icon">🏗️</div>
        <div class="exec-stat-value">{_exec_blocked_assemblies:,}</div>
        <div class="exec-stat-label">הרכבות מושפעות</div>
    </div>
    <div class="exec-stat">
        <div class="exec-stat-icon">🏭</div>
        <div class="exec-stat-value">{_exec_active_wip:,}</div>
        <div class="exec-stat-label">הרכבות ב-WIP פעיל</div>
    </div>
</div>
""", unsafe_allow_html=True)

NAV_GROUPS = {
    "📊 ניתוח ומעקב": [
        "📈 Executive Dashboard",
        "📊 תוכנית ייצור (Smart CTB)",
        "🏆 קיבולת ייצור מקסימלית",
        "📌 לוח סטטוסים (Kanban)",
        "📅 מעקב ETA ודחיות",
    ],
    "🛒 רכש": [
        "🛒 דשבורד רכש חוצה-פרויקטים",
    ],
    "⚙️ פעולות וניהול": [
        "💡 סימולציית What-If",
        "🏭 ניהול WIP (בייצור)",
        "📅 עדכון מלאי וספקים",
        "✏️ עריכת ETA מרוכזת",
    ],
    "🛠️ מנהלה": [
        "↩️ ניהול UNDO",
        "📦 ניהול מלאי מעודכן",
        "🎯 ניתוח רגישות ותוכנית",
        "ℹ️ הנחות עבודה לקובץ תקין",
    ],
}
PAGE_TO_GROUP = {page: g for g, items in NAV_GROUPS.items() for page in items}

if "main_nav_page" not in st.session_state:
    st.session_state["main_nav_page"] = "📈 Executive Dashboard"

st.markdown('<div class="nav-bar-label">🧭 ניווט מהיר — לחץ על קטגוריה לבחירת טאב</div>', unsafe_allow_html=True)
nav_cols = st.columns(len(NAV_GROUPS))
_has_popover = hasattr(st, "popover")
for (_group_label, _group_items), _col in zip(NAV_GROUPS.items(), nav_cols):
    with _col:
        _current_group = PAGE_TO_GROUP.get(st.session_state["main_nav_page"])
        _trigger_label = f"{_group_label}  ▾" + (f"    |    {st.session_state['main_nav_page']}" if _current_group == _group_label else "")
        _nav_container = st.popover(_trigger_label, use_container_width=True) if _has_popover else st.expander(_trigger_label, expanded=False)
        with _nav_container:
            for _item in _group_items:
                _btn_type = "primary" if st.session_state["main_nav_page"] == _item else "secondary"
                if st.button(_item, key=f"nav_btn_{_item}", use_container_width=True, type=_btn_type):
                    st.session_state["main_nav_page"] = _item
                    st.rerun()

nav_page = st.session_state["main_nav_page"]
st.markdown(f'<div class="section-title" style="font-size:22px;">{nav_page}</div>', unsafe_allow_html=True)

if nav_page == "📈 Executive Dashboard":
    israel_time = datetime.utcnow() + timedelta(hours=3)
    current_time_str = israel_time.strftime("%d/%m/%Y | %H:%M:%S")
    st.markdown(f"""
    <div style="display: flex; justify-content: space-between; align-items: center; margin-bottom: 15px; opacity: 0.85; font-weight: 600;">
        <div>🎯 תמונת מצב ניהולית לטווח חודשים: {', '.join(selected_target_yms)}</div>
        <div>🕒 שעון ישראל (עדכני): {current_time_str}</div>
    </div>
    """, unsafe_allow_html=True)

    dash_df = breakdown_df.copy()
    if selected_assembly != "הכל":
        dash_df = dash_df[dash_df["Assembly"] == selected_assembly]

    wip_cache_dash = fetch_wip_records(PROJECT_ID)
    inv_cache_dash = fetch_all_inventory_records(PROJECT_ID)

    total_planned_qty = 0.0
    total_executable_qty = 0.0
    total_planned_assemblies_count = 0
    blocked_assemblies = len(dash_df['Assembly'].unique()) if not dash_df.empty else 0

    assemblies_to_evaluate = [a for a in valid_assemblies if selected_assembly == "הכל" or a == selected_assembly]
    shared_plan = compute_shared_executable_plan(selected_target_yms, valid_assemblies, inv_cache_dash, wip_cache_dash)

    for asm_col in assemblies_to_evaluate:
        for target_m in selected_target_yms:
            info = shared_plan.get((asm_col, target_m))
            if info is None:
                continue
            total_planned_assemblies_count += 1
            total_planned_qty += info["raw_build"]
            total_executable_qty += info["net_executable"]

    readiness_pct = (total_executable_qty / total_planned_qty * 100) if total_planned_qty > 0 else 100
    unique_shortage_count = len(dash_df['PN'].unique()) if not dash_df.empty else 0
    active_wip_list = [(w, q) for w, q in wip_cache_dash.items() if q > 0]
    total_wip_active_count = len(active_wip_list)

    col_k1, col_k2, col_k3, col_k4, col_k5 = st.columns(5)
    with col_k1:
        kpi_card("🟢 מוכנות ייצור משוקללת", f"{readiness_pct:.1f}%", f"{total_executable_qty:,.0f} / {total_planned_qty:,.0f} יחידות ניתן לייצור", "green")
    with col_k2:
        kpi_card("🔴 הרכבות חסומות", blocked_assemblies, "בטווח הנבחר", "red")
    with col_k3:
        kpi_card("🏭 פעילים ב-WIP", total_wip_active_count, "הודעות ייצור פעילות", "blue")
    with col_k4:
        kpi_card("📦 מק'טים בגירעון", unique_shortage_count, "פריטים ייחודיים", "orange")
    with col_k5:
        kpi_card("📊 גירעון מצטברת", f"{dash_df['Total_MRP_Shortage'].sum():,.0f}" if not dash_df.empty else "0", "יחידות", "blue")

    with st.expander("🔍 הצג פירוט כרטיסי הרכבות פעילים ב-WIP (לחץ לפתיחה)", expanded=False):
        if active_wip_list:
            wip_detail_rows = [{"קוד הרכבה": asm_pn, "תיאור הרכבה": ASSEMBLY_BOM_TREE.get(asm_pn, {}).get("desc", ""), "כמות ב-WIP": asm_qty, "רמה בעץ": assembly_levels.get(asm_pn, 0)} for asm_pn, asm_qty in active_wip_list]
            st.dataframe(pd.DataFrame(wip_detail_rows), use_container_width=True)
        else:
            st.info("אין כרגע הרכבות פעילות ב-WIP.")

    st.divider()

    if not dash_df.empty and len(dash_df) > 0:
        col_g0, col_g1, col_g2 = st.columns([1, 1.2, 1.2])

        with col_g0:
            st.markdown("##### 🎯 מד מוכנות ייצור")
            fig_gauge = go.Figure(go.Indicator(
                mode="gauge+number",
                value=readiness_pct,
                number={'suffix': "%", 'font': {'size': 34}},
                gauge={
                    'axis': {'range': [0, 100]},
                    'bar': {'color': PRIMARY},
                    'steps': [
                        {'range': [0, 50], 'color': '#3B1F1F'},
                        {'range': [50, 80], 'color': '#3B2F1F'},
                        {'range': [80, 100], 'color': '#1F3B2A'},
                    ],
                }
            ))
            fig_gauge.update_layout(template=PLOTLY_TEMPLATE, height=260, margin=dict(t=10, b=10, l=20, r=20))
            st.plotly_chart(fig_gauge, use_container_width=True)

        with col_g1:
            st.markdown("##### 🥧 התפלגות חוסרים לפי סוג פריט")
            dash_df_unique_pn = dash_df.drop_duplicates(subset=["PN"])
            fig_pie = px.pie(dash_df_unique_pn, names="Item_Type", values="Total_MRP_Shortage", hole=0.5, color_discrete_sequence=COLOR_SEQ)
            fig_pie.update_layout(template=PLOTLY_TEMPLATE, height=280, margin=dict(t=10, b=10, l=10, r=10))
            st.plotly_chart(fig_pie, use_container_width=True)

        with col_g2:
            st.markdown("##### 💰 TOP 10 חוסרים לפי ערך כספי ($ , מחיר יחידה × כמות חסרה)")
            top10_value_df = dash_df_unique_pn[dash_df_unique_pn["Shortage_Value"] > 0].sort_values("Shortage_Value", ascending=False).head(10)
            if not top10_value_df.empty:
                top10_value_df = top10_value_df.assign(
                    label=top10_value_df.apply(lambda r: f"{r['PN']} ({str(r['Supplier'])[:14]})", axis=1)
                )
                fig_top10 = go.Figure(go.Funnel(
                    y=top10_value_df["label"],
                    x=top10_value_df["Shortage_Value"],
                    textposition="inside",
                    texttemplate="$%{value:,.0f}",
                    marker={"color": top10_value_df["Shortage_Value"], "colorscale": "Reds"},
                    connector={"line": {"color": PRIMARY, "width": 1}},
                    hovertemplate="<b>%{y}</b><br>ערך חוסר: $%{value:,.0f}<extra></extra>"
                ))
                fig_top10.update_layout(template=PLOTLY_TEMPLATE, height=280, margin=dict(t=10, b=10, l=10, r=10))
                st.plotly_chart(fig_top10, use_container_width=True)
            else:
                st.info("אין נתוני מחיר זמינים לחישוב ערך כספי של החוסרים בטווח הנוכחי.")

        st.markdown("##### 🏭 התפלגות חוסרים לפי ספק")
        supplier_agg = dash_df_unique_pn.groupby("Supplier", as_index=False).agg(
            Total_Shortage=("Total_MRP_Shortage", "sum"),
            Total_Value=("Shortage_Value", "sum"),
            Item_Count=("PN", "nunique")
        ).sort_values("Total_Value", ascending=False)
        fig_sup = px.bar(
            supplier_agg, x="Supplier", y="Total_Value", color="Supplier",
            color_discrete_sequence=COLOR_SEQ, hover_data=["Item_Count", "Total_Shortage"],
            labels={"Total_Value": "סה\"כ ערך חוסר ($)", "Supplier": "ספק"}
        )
        fig_sup.update_layout(template=PLOTLY_TEMPLATE, height=320, margin=dict(t=10, b=10, l=10, r=10), showlegend=False, xaxis_tickangle=-30)
        st.plotly_chart(fig_sup, use_container_width=True)

        st.markdown('<div class="section-title">📋 טבלת פירוט ניהולית עם אפשרות ייצוא וקישורי חיפוש מלאי</div>', unsafe_allow_html=True)
        display_df = dash_df[[
            "PN", "Description", "Item_Type", "Supplier", "Status", "Assembly", "Assembly_Desc",
            "Qty_Per_Assembly", "Assembly_Monthly_Build", "Required_Demand", "Stock", "Total_MRP_Shortage",
            "Unit_Price", "Shortage_Value", "AW_Data",
            "חיפוש במאוזר", "חיפוש בדיגיקי", "חיפוש ב-Findchips"
        ]].rename(columns={
            "PN": "מק'ט", "Description": "תיאור פריט", "Item_Type": "סוג פריט", "Supplier": "ספק",
            "Status": "סטטוס טיפול", "Assembly": "קוד הרכבה", "Assembly_Desc": "תיאור הרכבה",
            "Qty_Per_Assembly": "כמות נדרשת", "Assembly_Monthly_Build": "ת. ייצור",
            "Required_Demand": "ביקוש מדויק", "Stock": "מלאי", "Total_MRP_Shortage": "סך חוסר",
            "Unit_Price": "מחיר יחידה ($)", "Shortage_Value": "ערך חוסר ($)", "AW_Data": "עמודה AW"
        })

        def _shortage_color(val, vmax):
            if vmax <= 0:
                return ""
            ratio = min(1.0, float(val) / vmax)
            return f"background-color: rgba(239,{int(180 - ratio * 140)},{int(120 - ratio * 100)},0.55); color: white;"

        def _value_color(val, vmax):
            if vmax <= 0:
                return ""
            ratio = min(1.0, float(val) / vmax)
            return f"background-color: rgba(79,{int(120 - ratio * 60)},229,{0.15 + ratio * 0.5}); color: white;"

        sorted_display_df = display_df.sort_values(by="סך חוסר", ascending=False)
        max_shortage = sorted_display_df["סך חוסר"].max() if not sorted_display_df.empty else 0
        max_value = sorted_display_df["ערך חוסר ($)"].max() if not sorted_display_df.empty else 0

        styled = sorted_display_df.style.map(lambda v: _shortage_color(v, max_shortage), subset=["סך חוסר"]) \
                                       .map(lambda v: _value_color(v, max_value), subset=["ערך חוסר ($)"]) \
                                       .format({"ערך חוסר ($)": "${:,.0f}", "מחיר יחידה ($)": "${:,.2f}"})
        st.dataframe(styled, column_config={
            "חיפוש במאוזר": st.column_config.LinkColumn("🔗 מאוזר", display_text="פתח במאוזר"),
            "חיפוש בדיגיקי": st.column_config.LinkColumn("🔗 דיגיקי", display_text="פתח בדיגיקי"),
            "חיפוש ב-Findchips": st.column_config.LinkColumn("🔗 Findchips", display_text="פתח ב-Findchips")
        }, use_container_width=True)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            display_df.to_excel(writer, index=False, sheet_name='Executive_Shortages')
        st.download_button(label="📥 הורד דו'ח מנהלים מלא ל-Excel", data=output.getvalue(), file_name=f"MRP_Executive_Report_{selected_ym}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.success("🎉 אין חוסרים ב-MRP עבור ההגדרות והסינונים שנבחרו!")

elif nav_page == "📊 תוכנית ייצור (Smart CTB)":
    st.markdown(f'<div class="section-title">📊 סימולציית Clear To Build (CTB) מטריציונית עם השוואת כמויות וגרף הרכבות מפורט</div>', unsafe_allow_html=True)
    inv_cache_ctb = fetch_all_inventory_records(PROJECT_ID)
    wip_cache_ctb = fetch_wip_records(PROJECT_ID)
    matrix_rows, chart_assembly_data = [], []
    assemblies_to_check = [asm for asm in valid_assemblies if selected_assembly == "הכל" or asm == selected_assembly]

    shared_plan_ctb = compute_shared_executable_plan(selected_target_yms, valid_assemblies, inv_cache_ctb, wip_cache_ctb)

    for asm_col in assemblies_to_check:
        asm_desc = ASSEMBLY_BOM_TREE.get(asm_col, {}).get("desc", "")
        row_data = {"קוד הרכבה": asm_col, "תיאור הרכבה": asm_desc, "רמה בעץ": assembly_levels.get(asm_col, 0)}
        has_any_build = False

        for target_m in selected_target_yms:
            info = shared_plan_ctb.get((asm_col, target_m))
            raw_build = info["raw_build"] if info else 0.0
            current_wip_qty = info["wip"] if info else wip_cache_ctb.get(asm_col, 0.0)
            net_executable_qty = info["net_executable"] if info else 0.0

            if raw_build > 0 or current_wip_qty > 0:
                has_any_build = True

            row_data[f"תכנית ייצור ({target_m})"] = raw_build
            row_data[f"ניתן לייצור ({target_m})"] = net_executable_qty
            row_data[f"WIP ({target_m})"] = current_wip_qty

            if raw_build > 0 or current_wip_qty > 0:
                chart_assembly_data.append({"הרכבה ותיאור": f"{asm_col} - {asm_desc}", "חודש": target_m, "מדד": "תכנית ייצור", "כמות": raw_build})
                chart_assembly_data.append({"הרכבה ותיאור": f"{asm_col} - {asm_desc}", "חודש": target_m, "מדד": "ניתן לייצור בפועל", "כמות": net_executable_qty})
                chart_assembly_data.append({"הרכבה ותיאור": f"{asm_col} - {asm_desc}", "חודש": target_m, "מדד": "WIP", "כמות": current_wip_qty})

        for target_m in selected_target_yms:
            info = shared_plan_ctb.get((asm_col, target_m))
            raw_build = info["raw_build"] if info else 0.0
            gross_executable = info["gross_executable"] if info else 0.0
            current_wip_this_m = info["wip"] if info else 0.0
            limiting_pns = info["limiting_components"] if info else []
            remaining_target = max(0.0, raw_build - current_wip_this_m)

            if limiting_pns and gross_executable < remaining_target - 1e-9:
                formatted_missing = []
                for c_pn in limiting_pns:
                    c_match = df[df[PN_COL].astype(str).str.strip() == c_pn]
                    c_desc = str(c_match.iloc[0][DESC_COL]) if not c_match.empty else ""
                    raw_eta = get_first_supply_eta(c_pn, inv_cache_ctb)
                    shortage_amt = remaining_target - gross_executable
                    formatted_missing.append(f"{c_pn} ({c_desc[:10]}) - חוסם ל-{shortage_amt:g} יח' [ETA: {raw_eta}]")
                row_data[f"סטטוס וחוסרים ({target_m})"] = "❌ חסר (כולל התחשבות במלאי משותף עם הרכבות אחרות): " + " | ".join(formatted_missing)
            elif raw_build <= 0 and current_wip_this_m <= 0:
                row_data[f"סטטוס וחוסרים ({target_m})"] = "💤 ללא תוכנית ייצור"
            elif remaining_target <= 1e-9 and current_wip_this_m > 0:
                row_data[f"סטטוס וחוסרים ({target_m})"] = f"🔵 כל התוכנית ({raw_build:g} יח') כבר ב-WIP - אין יחידות נוספות לייצור החודש"
            else:
                row_data[f"סטטוס וחוסרים ({target_m})"] = "✅ מוכן לייצור מלא (כולל כל היחידות הנוספות שנותרו מעבר ל-WIP)"

        if has_any_build:
            matrix_rows.append(row_data)

    if matrix_rows:
        st.dataframe(pd.DataFrame(matrix_rows), use_container_width=True, height=420)
        if chart_assembly_data:
            fig_bar_asm = px.bar(pd.DataFrame(chart_assembly_data), x="הרכבה ותיאור", y="כמות", color="מדד", barmode="group", color_discrete_sequence=[PRIMARY, SUCCESS, ACCENT])
            fig_bar_asm.update_layout(template=PLOTLY_TEMPLATE, height=400, margin=dict(t=20, b=40, l=20, r=20), xaxis_tickangle=-25)
            st.plotly_chart(fig_bar_asm, use_container_width=True)

elif nav_page == "💡 סימולציית What-If":
    st.markdown('<div class="section-title">💡 סימולציית What-If (מה יקרה אם...)</div>', unsafe_allow_html=True)
    col_w1, col_w2 = st.columns(2)
    with col_w1:
        sim_pn = st.selectbox("בחר מק'ט לסימולציה", sorted(df[PN_COL].dropna().astype(str).unique()), key="sim_pn")
    with col_w2:
        sim_extra_stock = st.number_input("תוספת כמות מדומיינת למלאי", min_value=0.0, value=10.0, step=1.0)

    if st.button("🔮 הרץ סימולציית שחרור צוואר בקבוק"):
        sim_df = calculate_mrp_breakdown({sim_pn: sim_extra_stock}, target_yms=selected_target_yms)
        orig_blocked = set(breakdown_df['Assembly'].unique()) if not breakdown_df.empty else set()
        sim_blocked = set(sim_df['Assembly'].unique()) if not sim_df.empty else set()
        st.success(f"סימולציה הופעלה בהצלחה עבור מק'ט `{sim_pn}`.")
        col_m1, col_m2, col_m3 = st.columns(3)
        with col_m1:
            kpi_card("🟢 הרכבות שהשתחררו", len(orig_blocked - sim_blocked), "", "green")
        with col_m2:
            kpi_card("🔴 עדיין חסום", len(sim_blocked), "", "red")
        with col_m3:
            before_after_delta = (breakdown_df['Total_MRP_Shortage'].sum() if not breakdown_df.empty else 0) - (sim_df['Total_MRP_Shortage'].sum() if not sim_df.empty else 0)
            kpi_card("📉 צמצום גירעון", f"{before_after_delta:,.0f}", "יחידות", "blue")

elif nav_page == "📌 לוח סטטוסים (Kanban)":
    st.markdown('<div class="section-title">📌 לוח מעקב סטטוסים (Kanban Pipeline)</div>', unsafe_allow_html=True)
    statuses = [("פתוח", "📝 פתוח לטיפול", "#3B1F1F", DANGER), ("הוזמן", "🛒 הוזמן / בטיפול רכש", "#3B2F1F", WARNING), ("בדרך", "🚚 בדרך לקו", "#1F2A3B", ACCENT), ("התקבל", "✅ התקבל / סגור", "#1F3B2A", SUCCESS)]
    dedup_all = breakdown_df.drop_duplicates(subset=["PN"]) if not breakdown_df.empty else pd.DataFrame()
    kcols = st.columns(4)
    for (status_key, title, bg, accent_color), kcol in zip(statuses, kcols):
        with kcol:
            items = dedup_all[dedup_all["Status"] == status_key] if not dedup_all.empty else pd.DataFrame()
            st.markdown(f'<div class="kanban-col-header" style="background:{bg}; color:{accent_color};">{title} ({len(items)})</div>', unsafe_allow_html=True)
            for _, r in items.head(6).iterrows():
                st.markdown(f'<div class="kanban-card" style="border-color:{accent_color};"><b>{r["PN"]}</b><br><span style="opacity:0.75;">{str(r["Description"])[:24]}</span></div>', unsafe_allow_html=True)

elif nav_page == "🏭 ניהול WIP (בייצור)":
    st.markdown(f'<div class="section-title">🏭 ניהול WIP חכם (כולל סגירת מחזור ייצור ואימות היררכיה)</div>', unsafe_allow_html=True)
    wip_current = fetch_wip_records(PROJECT_ID)
    if wip_current:
        with st.form("close_wip_form"):
            wip_to_close = st.selectbox("בחר הרכבה שסיימה ייצור לחודש זה", list(wip_current.keys()), format_func=lambda x: f"{x} (כמות ב-WIP: {wip_current[x]})")
            is_finished = st.checkbox("האם ההרכבה הסתיימה לחלוטין והושלמה בהצלחה?")
            if st.form_submit_button("סגור WIP והוסף למלאי הזמין"):
                if is_finished:
                    closing_qty = wip_current[wip_to_close]
                    curr_stk, curr_eta, curr_stat, curr_sup, curr_comm, _, _ = get_inventory_record(wip_to_close)
                    save_inventory_record(wip_to_close, curr_stk + closing_qty, curr_eta, "התקבל", curr_sup, f"{curr_comm} | הושלם מייצור WIP בסך {closing_qty}", "WIP Close", webhook_url)
                    delete_wip_record(wip_to_close)
                    st.rerun()

        st.divider()
        st.markdown("##### 🗑️ מחיקת הרכבה מה-WIP (ללא העברה למלאי - לביטול/תיקון טעות)")
        col_del1, col_del2 = st.columns([3, 1])
        with col_del1:
            wip_to_delete = st.multiselect(
                "בחר הרכבה/ות למחיקה מה-WIP",
                list(wip_current.keys()),
                format_func=lambda x: f"{x} (כמות ב-WIP: {wip_current[x]})",
                key="wip_to_delete"
            )
        with col_del2:
            confirm_delete_wip = st.checkbox("מאשר מחיקה", key="confirm_delete_wip")
        if st.button("🗑️ מחק מה-WIP", key="delete_wip_btn"):
            if not wip_to_delete:
                st.warning("יש לבחור לפחות הרכבה אחת למחיקה.")
            elif not confirm_delete_wip:
                st.warning("יש לסמן 'מאשר מחיקה' לפני שמחיקה מתבצעת.")
            else:
                for asm_pn in wip_to_delete:
                    delete_wip_record(asm_pn)
                st.success(f"נמחקו {len(wip_to_delete)} הרכבות מה-WIP.")
                st.rerun()

    st.divider()
    st.markdown("##### ➕ צירוף הרכבה חדשה ל-WIP")
    wip_asm_choice = st.selectbox("בחר הרכבה חדשה לצירוף ל-WIP", filtered_assembly_cols, format_func=lambda x: assembly_mapping.get(x, x), key="wip_asm_choice")
    wip_qty_input = st.number_input("כמות יחידות הרכבה להוספה לייצור (WIP)", min_value=0.0, value=1.0, step=1.0, key="wip_qty_input")
    wip_target_month_label = st.selectbox(
        "לאיזה חודש בונים? (הבדיקה תתחשב באספקה/ETA שכבר אמורים להגיע עד חודש זה, לא רק במלאי הנוכחי בקופה)",
        list(month_options.keys()),
        index=list(month_options.keys()).index(selected_month_label) if selected_month_label in month_options else 0,
        key="wip_target_month_label"
    )
    wip_target_ym = pd.to_datetime(month_options[wip_target_month_label]).strftime("%Y-%m")

    hierarchy_blockers = check_hierarchical_ctb(wip_asm_choice, wip_qty_input, wip_target_ym) if wip_qty_input > 0 else []

    if hierarchy_blockers:
        st.error(f"⛔ נמצאו {len(hierarchy_blockers)} חוסרים בעץ ההרכבה עד חודש {wip_target_ym} (בהרכבה עצמה ו/או בתתי-ההרכבות שלה) - כולל התחשבות ב-ETA צפוי:")
        blockers_df = pd.DataFrame(hierarchy_blockers).rename(columns={
            "assembly": "קוד הרכבה חוסמת", "assembly_desc": "תיאור הרכבה חוסמת",
            "component": "מק\"ט חסר", "component_desc": "תיאור פריט",
            "required": "נדרש", "available": "זמין", "shortage": "חוסר"
        })
        st.dataframe(blockers_df, use_container_width=True, height=min(300, 45 + 35 * len(blockers_df)))
        with st.form("wip_form"):
            override_confirm = st.checkbox("⚠️ ידוע לי שיש חוסרים בעץ ההרכבה, ואני מאשר בכל זאת להוסיף ל-WIP (למשל אם מדובר בהזמנת ייצור מתוכננת מראש)")
            if st.form_submit_button("שמור ל-WIP בכל זאת"):
                if override_confirm:
                    save_wip_record(wip_asm_choice, wip_qty_input)
                    st.success("ההרכבה נוספה ל-WIP (עם חוסרים ידועים).")
                    st.rerun()
                else:
                    st.warning("יש לסמן את תיבת האישור כדי לשמור למרות החוסרים.")
    else:
        st.success("✅ נבדק עץ ההרכבה המלא - כל רכיבי הגלם וכל תתי-ההרכבות זמינים לכמות המבוקשת.")
        with st.form("wip_form"):
            if st.form_submit_button("בדיקת זמינות היררכית מלאה ושמור WIP"):
                save_wip_record(wip_asm_choice, wip_qty_input)
                st.success("ההרכבה נוספה בהצלחה ל-WIP!")
                st.rerun()

elif nav_page == "📅 עדכון מלאי וספקים":
    st.markdown('<div class="section-title">📅 עדכון מלאי, סטטוס ודחיית ספקים (ETA)</div>', unsafe_allow_html=True)
    selected_pn = search_pn if search_pn != "הכל" else st.selectbox("בחר מק'ט מכלל הפריטים לעדכון", sorted(df[PN_COL].dropna().astype(str).unique()), key="update_pn_select")
    if selected_pn != "הכל":
        saved_stock, saved_eta, saved_status, saved_supplier, saved_comment, saved_by, _ = get_inventory_record(selected_pn)
        _pn_match = df[df[PN_COL].astype(str).str.strip() == str(selected_pn).strip()]
        file_price = safe_num(_pn_match.iloc[0][PRICE_COL]) if not _pn_match.empty and PRICE_COL in df.columns else 0.0
        effective_price = get_effective_price(selected_pn, file_price)
        with st.form("inventory_form"):
            col_f1, col_f2, col_f3 = st.columns(3)
            with col_f1:
                added_stock_input = st.number_input("תוספת למלאי זמין", min_value=0.0, value=float(saved_stock), step=1.0)
            with col_f2:
                try: parsed_eta = pd.to_datetime(saved_eta).date() if saved_eta else date.today()
                except: parsed_eta = date.today()
                eta_date = st.date_input("תאריך הגעה מעודכן (ETA)", value=parsed_eta)
            with col_f3:
                status_options = ["פתוח", "הוזמן", "בייצור", "בדרך", "התקבל", "חסום"]
                status = st.selectbox("סטטוס טיפול", status_options, index=status_options.index(saved_status) if saved_status in status_options else 0)
            col_f4, col_f5 = st.columns(2)
            with col_f4:
                supplier = st.selectbox("ספק", supplier_options, index=supplier_options.index(saved_supplier) if saved_supplier in supplier_options else 0)
            with col_f5:
                price_input = st.number_input(
                    "מחיר יחידה ($) - קובע רק אם משנים אותו כאן, אחרת נשאר מחיר הקובץ",
                    min_value=0.0, value=float(effective_price), step=0.01, format="%.4f"
                )
                st.caption(f"מחיר בקובץ המקורי: ${file_price:,.4f}")
            comment = st.text_area("הערות", value=saved_comment)
            if st.form_submit_button("שמור עדכון קבוע בענן"):
                manual_price = price_input if abs(price_input - file_price) > 1e-9 else None
                save_inventory_record(selected_pn, added_stock_input, str(eta_date), status, supplier, comment, "User", webhook_url, unit_price=manual_price)
                st.success("העדכון נשמר!")
                st.rerun()

elif nav_page == "📅 מעקב ETA ודחיות":
    st.markdown('<div class="section-title">📅 מעקב ETA, דחיות, כמויות וקישורים למפיצים</div>', unsafe_allow_html=True)
    inv_cache_all = fetch_all_inventory_records(PROJECT_ID)
    eta_table_rows = []

    for _, row in df.iterrows():
        p_num = str(row[PN_COL]).strip()
        if not p_num or p_num == 'nan':
            continue
        p_desc = str(row[DESC_COL])
        original_p_type = str(row[ITEM_TYPE_COL]) if ITEM_TYPE_COL in df.columns else ""
        p_type = get_effective_item_type(p_num, original_p_type, inv_cache_all)

        orig_eta = get_base_mrp_eta(p_num)
        orig_qty = get_base_mrp_qty(p_num)
         
        saved_rec = inv_cache_all.get(p_num, {})
        current_eta_raw = saved_rec.get("eta", "")
        current_added_stock = saved_rec.get("added_stock", 0.0)
        curr_eta_fmt = pd.to_datetime(current_eta_raw).strftime("%Y-%m") if current_eta_raw else orig_eta

        eta_table_rows.append({
            "מק'ט": p_num,
            "תיאור פריט": p_desc,
            "סוג פריט": p_type,
            "ETA מקורי (MRP)": orig_eta,
            "כמות מקורית": orig_qty,
            "ETA מעודכן": curr_eta_fmt,
            "כמות מעודכנת": current_added_stock,
            "ספק": get_effective_supplier(p_num, inv_cache_all),
            "חיפוש במאוזר": f"https://www.mouser.co.il/c/?q={p_num}",
            "חיפוש בדיגיקי": f"https://www.digikey.com/en/products/result?keywords={p_num}",
            "חיפוש ב-Findchips": f"https://www.findchips.com/search/{p_num}"
        })

    eta_df = pd.DataFrame(eta_table_rows)
    if not eta_df.empty:
        needs_review_count = int((eta_df["ETA מקורי (MRP)"] == "בדיקה נדרשת").sum())
        st.metric("🔎 מק\"טים ללא ETA בסיסי במערכת (דורשים בדיקה ידנית)", needs_review_count)

        def _highlight_needs_review(val):
            if val == "בדיקה נדרשת":
                return f"background-color: {WARNING}; color: white; font-weight: 700;"
            return ""

        eta_df_styled = eta_df.style.map(_highlight_needs_review, subset=["ETA מקורי (MRP)"])
        st.dataframe(eta_df_styled, column_config={
            "חיפוש במאוזר": st.column_config.LinkColumn("🔗 מאוזר", display_text="פתח במאוזר"),
            "חיפוש בדיגיקי": st.column_config.LinkColumn("🔗 דיגיקי", display_text="פתח בדיגיקי"),
            "חיפוש ב-Findchips": st.column_config.LinkColumn("🔗 Findchips", display_text="פתח ב-Findchips")
        }, use_container_width=True, height=450)

elif nav_page == "↩️ ניהול UNDO":
    st.markdown('<div class="section-title">↩️ חזרה לאחור וניהול היסטוריה (UNDO)</div>', unsafe_allow_html=True)
    try: updated_items = supabase.table("mrp_inventory_updates").select("*").eq("project_id", PROJECT_ID).order("updated_at", desc=True).execute().data or []
    except: updated_items = []
    for item in updated_items:
        col_u1, col_u2, col_u3 = st.columns([3, 4, 1])
        with col_u1: st.markdown(f"**מק'ט:** `{item.get('pn')}`")
        with col_u2: st.text(f"תוספת: {item.get('added_stock')} | ETA: {item.get('eta')}")
        with col_u3:
            if st.button("🔄 UNDO", key=f"undo_{item.get('pn')}"):
                delete_inventory_record(item.get('pn'))
                st.rerun()

elif nav_page == "📦 ניהול מלאי מעודכן":
    st.markdown('<div class="section-title">📦 ניהול מלאי מעודכן (עריכה וגריעת כמויות)</div>', unsafe_allow_html=True)
    active_stock_items = {k: v for k, v in fetch_all_inventory_records(PROJECT_ID).items() if float(v.get("added_stock", 0.0)) > 0}
    if active_stock_items:
        st.dataframe(pd.DataFrame([{"מק'ט": k, "כמות": v["added_stock"], "ETA": v["eta"]} for k, v in active_stock_items.items()]), use_container_width=True)
        selected_mgmt_pn = st.selectbox("בחר מק'ט לעריכה או גריעה", list(active_stock_items.keys()), key="mgmt_pn_select")
        if selected_mgmt_pn:
            with st.form("edit_mgmt_form"):
                new_qty = st.number_input("עדכן כמות", min_value=0.0, value=float(active_stock_items[selected_mgmt_pn]["added_stock"]), step=1.0)
                if st.form_submit_button("🗑️ אפס או עדכן"):
                    delete_inventory_record(selected_mgmt_pn)
                    if new_qty > 0:
                        save_inventory_record(selected_mgmt_pn, new_qty, "", "פתוח", "אופק", "", "Tab 9", webhook_url)
                    st.rerun()

elif nav_page == "🎯 ניתוח רגישות ותוכנית":
    st.markdown('<div class="section-title">🎯 ניתוח רגישות וניהול תוכנית הייצור (עריכה פרטנית לפי הרכבה וחודש)</div>', unsafe_allow_html=True)
    if not assembly_plan_df.empty:
        orig_pivot_plan = assembly_plan_df.pivot_table(index=["Assembly_PN"], columns="YearMonth", values="Build_Qty", fill_value=0.0).reset_index()
        orig_pivot_plan.insert(1, "רמה", orig_pivot_plan["Assembly_PN"].map(lambda x: assembly_levels.get(x, 0)))
        orig_pivot_plan.insert(2, "תיאור הרכבה", orig_pivot_plan["Assembly_PN"].map(lambda x: assembly_mapping.get(x, x)))
        orig_pivot_plan = orig_pivot_plan.sort_values(by=["רמה", "Assembly_PN"])
        st.dataframe(orig_pivot_plan, use_container_width=True, height=280)

    st.divider()
    st.markdown("##### ⚙️ הגדרת שינוי רגישות: גורף או חודש ספציפי")
     
    col_mode_choice = st.columns(2)
    with col_mode_choice[0]:
        sens_scope = st.radio("היקף השינוי", ["שינוי גורף לכל החודשים", "שינוי לחודש ספציפי בלבד"], horizontal=True, key="sens_scope")
    with col_mode_choice[1]:
        if sens_scope == "שינוי לחודש ספציפי בלבד":
            available_yms = sorted(assembly_plan_df["YearMonth"].unique())
            target_sens_month = st.selectbox("בחר חודש ספציפי לעדכון", available_yms, key="target_sens_month")

    col_sens1, col_sens2, col_sens3 = st.columns([1.2, 1, 1])
    with col_sens1:
        sens_assembly_target = st.selectbox("בחר הרכבה לניתוח רגישות", ["הכל (כלל ההרכבות)"] + filtered_assembly_cols, format_func=lambda x: assembly_mapping.get(x, x), key="sens_assembly_target")
    with col_sens2:
        sens_mode = st.radio("סוג שינוי", ["אחוזים (%)", "מספרי (יחידות)"], horizontal=True, key="sens_mode")
    with col_sens3:
        if sens_mode == "אחוזים (%)":
            sensitivity_val = st.slider("שינוי אחוז תוכנית הייצור (%)", -50, 100, 0, 5, key="sens_slider")
        else:
            sensitivity_val = st.number_input("תוספת/הפחתה מספרית (יחידות)", -500, 500, 0, 1, key="sens_num")

    if st.button("🚀 הרץ ניתוח רגישות לתוכנית", key="run_sensitivity"):
        simulated_plan_df = assembly_plan_df.copy()
        sys_factor_map = ASSEMBLY_SYSTEM_FACTORS
         
        if sens_scope == "שינוי גורף לכל החודשים":
            if sens_mode == "אחוזים (%)" and sensitivity_val != 0:
                multiplier = 1.0 + (sensitivity_val / 100.0)
                if sens_assembly_target == "הכל (כלל ההרכבות)":
                    simulated_plan_df["Raw_Build_Qty"] *= multiplier
                    simulated_plan_df["Build_Qty"] *= multiplier
                else:
                    mask = simulated_plan_df["Assembly_PN"] == sens_assembly_target
                    simulated_plan_df.loc[mask, "Raw_Build_Qty"] *= multiplier
                    simulated_plan_df.loc[mask, "Build_Qty"] *= multiplier
            elif sens_mode == "מספרי (יחידות)" and sensitivity_val != 0:
                if sens_assembly_target == "הכל (כלל ההרכבות)":
                    for idx_row in simulated_plan_df.index:
                        asm_code = str(simulated_plan_df.loc[idx_row, "Assembly_PN"]).strip()
                        sys_f = sys_factor_map.get(asm_code, 1)
                        curr_raw = simulated_plan_df.loc[idx_row, "Raw_Build_Qty"]
                        new_raw = max(0.0, curr_raw + (sensitivity_val / sys_f))
                        simulated_plan_df.loc[idx_row, "Raw_Build_Qty"] = new_raw
                        simulated_plan_df.loc[idx_row, "Build_Qty"] = new_raw * sys_f
                else:
                    mask = simulated_plan_df["Assembly_PN"] == sens_assembly_target
                    sys_f = sys_factor_map.get(sens_assembly_target, 1)
                    for idx_row in simulated_plan_df[mask].index:
                        curr_raw = simulated_plan_df.loc[idx_row, "Raw_Build_Qty"]
                        new_raw = max(0.0, curr_raw + (sensitivity_val / sys_f))
                        simulated_plan_df.loc[idx_row, "Raw_Build_Qty"] = new_raw
                        simulated_plan_df.loc[idx_row, "Build_Qty"] = new_raw * sys_f
        else:
            if sens_mode == "אחוזים (%)" and sensitivity_val != 0:
                multiplier = 1.0 + (sensitivity_val / 100.0)
                if sens_assembly_target == "הכל (כלל ההרכבות)":
                    mask = simulated_plan_df["YearMonth"] == target_sens_month
                    simulated_plan_df.loc[mask, "Raw_Build_Qty"] *= multiplier
                    simulated_plan_df.loc[mask, "Build_Qty"] *= multiplier
                else:
                    mask = (simulated_plan_df["Assembly_PN"] == sens_assembly_target) & (simulated_plan_df["YearMonth"] == target_sens_month)
                    simulated_plan_df.loc[mask, "Raw_Build_Qty"] *= multiplier
                    simulated_plan_df.loc[mask, "Build_Qty"] *= multiplier
            elif sens_mode == "מספרי (יחידות)" and sensitivity_val != 0:
                if sens_assembly_target == "הכל (כלל ההרכבות)":
                    mask = simulated_plan_df["YearMonth"] == target_sens_month
                    for idx_row in simulated_plan_df[mask].index:
                        asm_code = str(simulated_plan_df.loc[idx_row, "Assembly_PN"]).strip()
                        sys_f = sys_factor_map.get(asm_code, 1)
                        curr_raw = simulated_plan_df.loc[idx_row, "Raw_Build_Qty"]
                        new_raw = max(0.0, curr_raw + (sensitivity_val / sys_f))
                        simulated_plan_df.loc[idx_row, "Raw_Build_Qty"] = new_raw
                        simulated_plan_df.loc[idx_row, "Build_Qty"] = new_raw * sys_f
                else:
                    mask = (simulated_plan_df["Assembly_PN"] == sens_assembly_target) & (simulated_plan_df["YearMonth"] == target_sens_month)
                    sys_f = sys_factor_map.get(sens_assembly_target, 1)
                    for idx_row in simulated_plan_df[mask].index:
                        curr_raw = simulated_plan_df.loc[idx_row, "Raw_Build_Qty"]
                        new_raw = max(0.0, curr_raw + (sensitivity_val / sys_f))
                        simulated_plan_df.loc[idx_row, "Raw_Build_Qty"] = new_raw
                        simulated_plan_df.loc[idx_row, "Build_Qty"] = new_raw * sys_f

        st.session_state["temp_simulated_plan"] = simulated_plan_df
        st.success("ניתוח הרגישות בוצע בהצלחה! צפה בתוצאות המעודכנות למטה.")

    if "temp_simulated_plan" in st.session_state:
        st.divider()
        st.markdown("##### 📋 תצוגה מקדימה של התוכנית הסימולטיבית (לאחר ניתוח רגישות):")
        preview_pivot = st.session_state["temp_simulated_plan"].pivot_table(index=["Assembly_PN"], columns="YearMonth", values="Build_Qty", fill_value=0.0).reset_index()
        preview_pivot.insert(1, "רמה", preview_pivot["Assembly_PN"].map(lambda x: assembly_levels.get(x, 0)))
        preview_pivot.insert(2, "תיאור הרכבה", preview_pivot["Assembly_PN"].map(lambda x: assembly_mapping.get(x, x)))
        preview_pivot = preview_pivot.sort_values(by=["רמה", "Assembly_PN"])
        st.dataframe(preview_pivot, use_container_width=True, height=240)

        with st.form("update_plan_form"):
            update_confirmation = st.checkbox("❓ האם אתה מאשר לשמור את השינויים ולהחיל את תוכנית הייצור החדשה על כלל המערכת?")
            if st.form_submit_button("💾 שמור שינויים ועדכן את תוכנית העבודה"):
                if update_confirmation:
                    st.session_state["previous_approved_plan"] = assembly_plan_df.copy()
                    st.session_state["custom_assembly_plan_df"] = st.session_state["temp_simulated_plan"]
                    save_cloud_assembly_plan(st.session_state["temp_simulated_plan"], PROJECT_ID)
                    del st.session_state["temp_simulated_plan"]
                    st.success("תוכנית הייצור עודכנה ונשמרה בהצלחה בענן ובמערכת!")
                    st.rerun()
                else:
                    st.warning("יש לסמן את תיבת האישור כדי לשמור את השינויים.")

    if "previous_approved_plan" in st.session_state:
        st.divider()
        st.markdown("##### 📊 השוואה בין תוכנית הייצור הקודמת לחדשה")
        orig_plan_pivot = st.session_state["previous_approved_plan"].pivot_table(index="Assembly_PN", columns="YearMonth", values="Build_Qty", fill_value=0.0)
        new_plan_pivot = assembly_plan_df.pivot_table(index="Assembly_PN", columns="YearMonth", values="Build_Qty", fill_value=0.0)
        comparison_diff = new_plan_pivot.sub(orig_plan_pivot, fill_value=0.0).reset_index()
        comparison_diff.insert(1, "רמה", comparison_diff["Assembly_PN"].map(lambda x: assembly_levels.get(x, 0)))
        comparison_diff.insert(2, "תיאור הרכבה", comparison_diff["Assembly_PN"].map(lambda x: assembly_mapping.get(x, x)))
        comparison_diff = comparison_diff.sort_values(by=["רמה", "Assembly_PN"])
        st.dataframe(comparison_diff, use_container_width=True)

elif nav_page == "✏️ עריכת ETA מרוכזת":
    st.markdown('<div class="section-title">✏️ עריכת ETA מרוכזת לכל הפריטים</div>', unsafe_allow_html=True)
    st.caption("עדכון כאן נשמר ישירות ב-DB, ומשפיע מיידית על כל חישובי ה-MRP, ה-CTB ובדיקת הזמינות ההיררכית ב-WIP בכל שאר הטאבים.")

    with st.expander(f"📥 ייבוא מרוכז של רשימת ספקים ({len(DEFAULT_SUPPLIER_MAP)} פריטים) - עדכון חד-פעמי ל-DB", expanded=False):
        st.caption("לחיצה על הכפתור תעדכן את שדה ה'ספק' ב-DB עבור כל הפריטים ברשימה, בלי לגעת בשדות אחרים (תוספת מלאי, ETA, סטטוס, הערות) שכבר קיימים לאותם פריטים.")
        inv_cache_supplier_preview = fetch_all_inventory_records(PROJECT_ID)
        already_matching = sum(
            1 for pn, sup in DEFAULT_SUPPLIER_MAP.items()
            if inv_cache_supplier_preview.get(pn, {}).get("supplier", "") == sup
        )
        st.write(f"מתוך {len(DEFAULT_SUPPLIER_MAP)} פריטים ברשימה: {already_matching} כבר תואמים ב-DB, {len(DEFAULT_SUPPLIER_MAP) - already_matching} ידרשו עדכון.")
        if st.button("💾 עדכן את כל רשימת הספקים ב-DB עכשיו", key="bulk_supplier_import_btn"):
            count, err = bulk_update_suppliers(DEFAULT_SUPPLIER_MAP, inv_cache_supplier_preview)
            if err:
                st.error(f"שגיאה בעדכון מרוכז: {err}")
            else:
                st.success(f"עודכנו {count} פריטים עם שיוך הספק שלהם ב-DB.")
                st.rerun()

    st.divider()
    inv_cache_bulk = fetch_all_inventory_records(PROJECT_ID)

    col_bf1, col_bf2, col_bf3 = st.columns([1.2, 1.5, 1])
    with col_bf1:
        bulk_item_type = st.selectbox("סינון לפי סוג פריט", ["הכל"] + item_types, key="bulk_item_type")
    with col_bf2:
        bulk_search = st.text_input("חיפוש לפי מק\"ט או תיאור", key="bulk_search")
    with col_bf3:
        bulk_only_shortage = st.checkbox("הצג רק פריטים שכרגע בחוסר", key="bulk_only_shortage")

    shortage_pns = set(breakdown_df["PN"].unique()) if not breakdown_df.empty else set()
    bulk_rows = []

    for _, row in df.iterrows():
        p_num = str(row[PN_COL]).strip()
        if not p_num or p_num == 'nan':
            continue
        if bulk_only_shortage and p_num not in shortage_pns:
            continue
        p_desc = str(row[DESC_COL])
        original_p_type = str(row[ITEM_TYPE_COL]) if ITEM_TYPE_COL in df.columns else ""
        p_type = get_effective_item_type(p_num, original_p_type, inv_cache_bulk)
        if bulk_item_type != "הכל" and p_type != bulk_item_type:
            continue
        if bulk_search and bulk_search.strip():
            needle = bulk_search.strip().lower()
            if needle not in p_num.lower() and needle not in p_desc.lower():
                continue

        orig_eta = get_base_mrp_eta(p_num)
        orig_qty = get_base_mrp_qty(p_num)
        saved_rec = inv_cache_bulk.get(p_num, {})
        current_eta_raw = saved_rec.get("eta", "")
        try:
            current_eta_date = pd.to_datetime(current_eta_raw).date() if current_eta_raw else None
        except Exception:
            current_eta_date = None

        bulk_rows.append({
            "מק\"ט": p_num,
            "תיאור פריט": p_desc,
            "סוג פריט": p_type,
            "ETA מקורי (MRP)": orig_eta,
            "כמות מקורית (MRP)": orig_qty,
            "ETA מעודכן": current_eta_date,
            "תוספת מלאי": float(saved_rec.get("added_stock", 0.0) or 0.0),
            "סטטוס": saved_rec.get("status", "פתוח") or "פתוח",
            "ספק": get_effective_supplier(p_num, inv_cache_bulk),
            "הערות": saved_rec.get("comment", "") or "",
        })

    if not bulk_rows:
        st.info("לא נמצאו פריטים התואמים לסינון שנבחר.")
    else:
        bulk_df = pd.DataFrame(bulk_rows)
        st.caption(f"מציג {len(bulk_df)} פריטים. ניתן לערוך ETA מעודכן / תוספת מלאי / סטטוס / ספק / הערות ישירות בטבלה, ואז ללחוץ על 'שמור' למטה.")

        edited_df = st.data_editor(
            bulk_df,
            key="bulk_eta_editor",
            use_container_width=True,
            height=520,
            hide_index=True,
            disabled=["מק\"ט", "תיאור פריט", "ETA מקורי (MRP)", "כמות מקורית (MRP)"],
            column_config={
                "ETA מעודכן": st.column_config.DateColumn("ETA מעודכן", format="YYYY-MM-DD"),
                "תוספת מלאי": st.column_config.NumberColumn("תוספת מלאי", min_value=0.0, step=1.0),
                "סטטוס": st.column_config.SelectboxColumn("סטטוס", options=["פתוח", "הוזמן", "בייצור", "בדרך", "התקבל", "חסום"]),
                "ספק": st.column_config.SelectboxColumn("ספק", options=supplier_options),
                "סוג פריט": st.column_config.SelectboxColumn("סוג פריט", options=item_types),
            }
        )

        if st.button("💾 שמור את כל השינויים ל-DB", key="bulk_save_btn"):
            changed_count = 0
            for i in range(len(bulk_df)):
                orig_row = bulk_df.iloc[i]
                new_row = edited_df.iloc[i]

                new_stock_val = float(new_row["תוספת מלאי"]) if pd.notnull(new_row["תוספת מלאי"]) else 0.0
                changed = (
                    str(orig_row["ETA מעודכן"]) != str(new_row["ETA מעודכן"]) or
                    float(orig_row["תוספת מלאי"]) != new_stock_val or
                    orig_row["סטטוס"] != new_row["סטטוס"] or
                    orig_row["ספק"] != new_row["ספק"] or
                    orig_row["הערות"] != new_row["הערות"] or
                    orig_row["סוג פריט"] != new_row["סוג פריט"]
                )
                if changed:
                    save_inventory_record(
                        pn=orig_row["מק\"ט"],
                        added_stock=new_stock_val,
                        eta=str(new_row["ETA מעודכן"]) if new_row["ETA מעודכן"] else "",
                        status=new_row["סטטוס"],
                        supplier=new_row["ספק"],
                        comment=new_row["הערות"],
                        updated_by="Bulk ETA Editor",
                        webhook_url=webhook_url,
                        item_type=new_row["סוג פריט"]
                    )
                    changed_count += 1
            if changed_count > 0:
                st.success(f"נשמרו {changed_count} שינויים ל-DB. כל החישובים בכל הטאבים יתעדכנו בהתאם.")
                st.rerun()
            else:
                st.info("לא זוהו שינויים לשמירה.")

elif nav_page == "🏆 קיבולת ייצור מקסימלית":
    st.markdown('<div class="section-title">🏆 קיבולת ייצור מקסימלית לכל הרכבה</div>', unsafe_allow_html=True)
    st.caption(
        "כמה יחידות מכל הרכבה ניתן היה לייצר תיאורטית עד חודש נתון, בהתבסס על מלאי + אספקה עם ETA שכבר חל + ניכוי מה שכבר נצרך ב-WIP - "
        "ללא קשר לתוכנית הייצור המתוכננת. חישוב זה מבוצע לכל הרכבה בנפרד (תקרת קיבולת עליונה), ולא מחלק מלאי משותף בין הרכבות כמו בטאב 'תוכנית ייצור'."
    )

    CAP_COL = "קיבולת מקסימלית (יחידות)"

    cap_target_month_label = st.selectbox(
        "עד איזה חודש לבדוק זמינות (כולל אספקה עם ETA עד לפני חודש זה)",
        list(month_options.keys()),
        index=list(month_options.keys()).index(selected_month_label) if selected_month_label in month_options else 0,
        key="cap_target_month_label"
    )
    cap_target_ym = pd.to_datetime(month_options[cap_target_month_label]).strftime("%Y-%m")

    inv_cache_cap = fetch_all_inventory_records(PROJECT_ID)
    wip_cache_cap = fetch_wip_records(PROJECT_ID)

    cap_rows = []
    for asm_col in valid_assemblies:
        max_qty, limiting = compute_max_buildable(asm_col, cap_target_ym, inv_cache_cap, wip_cache_cap)
        current_wip_qty = wip_cache_cap.get(asm_col, 0.0)
        cap_rows.append({
            "קוד הרכבה": asm_col,
            "תיאור": assembly_mapping.get(asm_col, asm_col),
            "רמה בעץ": assembly_levels.get(asm_col, 0),
            CAP_COL: round(max_qty, 1),
            "כבר ב-WIP (יחידות)": current_wip_qty,
            "רכיב/הרכבה מגבילים": limiting if limiting else "—"
        })

    cap_df = pd.DataFrame(cap_rows).sort_values(["רמה בעץ", CAP_COL], ascending=[True, False])

    col_cap1, col_cap2 = st.columns([1, 1])
    with col_cap1:
        if not cap_df.empty:
            top_row = cap_df.iloc[0]
            top_asm_name = top_row["קוד הרכבה"]
            top_asm_value = f"{top_row[CAP_COL]:,.0f} יחידות"
        else:
            top_asm_name = "—"
            top_asm_value = ""
        st.metric("🏆 ההרכבה עם הקיבולת הגבוהה ביותר", top_asm_name, top_asm_value)

    with col_cap2:
        if not cap_df.empty:
            bottleneck_row = cap_df.loc[cap_df[CAP_COL].idxmin()]
            bottleneck_asm_name = bottleneck_row["קוד הרכבה"]
            bottleneck_asm_value = f"{bottleneck_row[CAP_COL]:,.0f} יחידות"
        else:
            bottleneck_asm_name = "—"
            bottleneck_asm_value = ""
        st.metric("🔻 ההרכבה עם הקיבולת הנמוכה ביותר (צוואר בקבוק)", bottleneck_asm_name, bottleneck_asm_value)

    if not cap_df.empty:
        fig_cap = px.bar(
            cap_df.sort_values(CAP_COL, ascending=True),
            x=CAP_COL, y="קוד הרכבה", orientation='h',
            color=CAP_COL, color_continuous_scale="Blues",
            hover_data=["תיאור", "רכיב/הרכבה מגבילים"]
        )
        fig_cap.update_layout(template=PLOTLY_TEMPLATE, height=max(400, 28 * len(cap_df)), margin=dict(t=10, b=10, l=10, r=10))
        st.plotly_chart(fig_cap, use_container_width=True)
        st.dataframe(cap_df, use_container_width=True, height=420)
    else:
        st.info("אין הרכבות זמינות לחישוב קיבולת מקסימלית עבור החודש שנבחר.")

elif nav_page == "ℹ️ הנחות עבודה לקובץ תקין":
    st.markdown("""
מדריך זה נבנה בעקבות כמה קבצי פרויקטים שונים שנבדקו בפועל, כדי שתדע מראש מה קובץ Excel חדש
צריך להכיל כדי שהמערכת תוכל לקרוא אותו בלי שגיאות - ומה קורה כשמשהו חסר.
""")

    st.markdown("### 1️⃣ שורת כותרת ראשית")
    st.markdown(
        "חייבת להיות שורה אחת עם `#` בעמודה הראשונה ו-`PN_ID` באחת מ-5 העמודות הבאות אחריה "
        "(אפשר עמודה ריקה ביניהן). זו שורת הכותרת של הטבלה הראשית - כל מה שמימין לה (REV, "
        "CATALOG, DESCRIPTION וכו') והרכבות המטריצה."
    )

    st.markdown("### 2️⃣ מטריצת \"היכן משתמשים\" (where-used)")
    st.markdown(
        "אחרי עמודות המטא-דאטה, כל עמודה שהכותרת שלה היא מק\"ט הרכבה (למשל `1094T030-002`) "
        "נחשבת \"עמודת הרכבה\". כל שורה מתחתיה היא רכיב, והערך בתא הוא כמות הרכיב הזה ביחידה "
        "אחת מההרכבה. עמודה בשם `SUM` בסוף המטריצה (לא חובה, אבל עוזרת לזיהוי הגבול)."
    )

    st.markdown("### 3️⃣ עץ ה-BOM (מק\"ט/רמה/תיאור)")
    st.markdown(
        "אפשר בשתי דרכים - המערכת מנסה את שתיהן לפי הסדר:\n\n"
        "**דרך א' (מועדפת):** בלוק של 3 עמודות סמוכות עם הכותרות `DESC` (או `DESCRIPTION`), "
        "`LEVEL`, ו-`PN` (לא תלוי רישיות - `desc`/`level`/`pn` גם עובד). מתחתיהן, שורה אחת "
        "לכל פריט בעץ: תיאור, מספר רמה (0 = הכי עליון), ומק\"ט. אפשר גם עמודת כמות רביעית "
        "מיד אחרי `PN` אם השם שלה מכיל את המילה `QTY` - זו הכמות ליחידת-מערכת. **הבלוק הזה "
        "יכול לשבת בכל מקום בגיליון** (גם מעל שורת הכותרת הראשית) - המערכת סורקת את כל "
        "הגיליון, לא רק מתחת לכותרת.\n\n"
        "**דרך ב' (גיבוי, אם אין בלוק כזה בכלל):** שתי \"שורות תיוג\" ממש מעל שורת הכותרת "
        "הראשית - אחת עם `DESC`/`DESCRIPTION` ואחת עם `LEVEL` (סובלני גם לשגיאות כתיב כמו "
        "`LEVAL`), שמתארות ישירות את עמודות מטריצת ה-where-used (רמה ותיאור לכל הרכבה, "
        "בעמודה המתאימה לה). במקרה הזה הכמות ליחידת-הורה נגזרת אוטומטית מהמטריצה עצמה."
    )
    st.warning(
        "⚠️ אם אף אחת מהדרכים לא נמצאת, המערכת עוצרת עם הודעת שגיאה ברורה - "
        "היא לא ממשיכה לרוץ עם נתונים חלקיים או שגויים."
    )

    st.markdown("### 4️⃣ בלוקי תאריכים (תוכנית עבודה / DEMAND / אספקה-ETA)")
    st.markdown(
        "אורך אופק התכנון **גמיש לחלוטין** - יכול להיות כמה חודשים בודדים או כמה שנים "
        "(נבדק עד כ-12 שנה קדימה). המערכת מזהה את בלוק \"תוכנית העבודה\" לפי המיקום שמיד "
        "אחרי עמודות עץ ה-BOM, ואת סופו לפי תווית `DEMAND` שנמצאת בגיליון (אם יש). בלוק "
        "האספקה (ETA) מזוהה כרצף התאריכים הרציף הארוך ביותר שמסתיים ממש לפני בלוק עץ ה-BOM."
    )

    st.markdown("### 5️⃣ עמודות אופציונליות (לא חובה, אבל בלעדיהן פיצ'רים מסוימים לא יהיו זמינים)")
    st.markdown(
        "- **`STOCK`** - מלאי בסיס לכל רכיב.\n"
        "- **`סיווג פריט`** (או `ITEM_TYPE`) - בלעדיה, סינון/עריכת סוג פריט לא יהיו זמינים.\n"
        "- **מחיר** (`PRICE_CALC`, `TARGET PRICE`, `TOTAL PRICE`, או `PRICE`) - בלעדיה, "
        "חישובי ערך כספי לא יהיו זמינים.\n"
        "- **ספק** (`ספק`, `SUPPLIER`, `MANUFACTURER`, או `POC SUPPLIER`)."
    )

    st.markdown("### ⚠️ מגבלה ידועה: הרכבה משותפת בין שני ענפים")
    st.markdown(
        "אם אותה הרכבה (אותו מק\"ט) מופיעה תחת **שני הורים שונים** (בשני ענפים נפרדים "
        "בעץ) - המערכת תומכת בזה בחישובי ה-CTB/קיבולת מקסימלית (לכל קשר הורה-ילד יש כמות "
        "משלו), אבל **בתצוגה** (למשל ברשימת ההרכבות, בתיאור, ברמה) יוצג רק ייצוג אחד "
        "\"כללי\" לפריט - לא שני ייצוגים נפרדים לפי ענף."
    )

    st.markdown("### 📎 גיליונות מרובים")
    st.markdown(
        "המערכת קוראת **רק את הגיליון הראשון** בקובץ. אם תוכנית העבודה/הביקוש יושבת "
        "בגיליון נפרד (לא הגיליון הראשי), היא לא תיקרא - הנתונים חייבים לשבת באותו "
        "גיליון כמו עץ ה-BOM והמטריצה."
    )

    st.markdown("### ➕ איך מוסיפים פרויקט חדש")
    st.markdown("**שלב 1 - להעלות את קובץ ה-Excel ל-GitHub ולהעתיק ממנו קישור raw:**")
    st.markdown(
        "1. תעלה את הקובץ לריפו ב-GitHub (לא משנה השם, ולא משנה אם זה אותו ריפו של קבצים "
        "קודמים או ריפו נפרד).\n"
        "2. תפתח את הקובץ בתוך GitHub ותלחץ על הכפתור **Raw** מעל תצוגת הקובץ. הדפדפן יפתח "
        "כתובת שמתחילה ב-`raw.githubusercontent.com`.\n"
        "3. תעתיק את הכתובת הזו משורת הדפדפן.\n\n"
        "**קיצור דרך אם אין לך גישה לכפתור:** קח את הכתובת הרגילה מהדפדפן (שמכילה `/blob/`), "
        "והדבק אותה כמו שהיא - האפליקציה ממירה אותה אוטומטית לקישור raw בעצמה, אין צורך "
        "לערוך אותה ידנית."
    )
    st.markdown("**שלב 2 - להוסיף אותו באפליקציה:**")
    st.markdown(
        "בסיידבר, תחת \"📁 פרויקט פעיל\", תבחר באפשרות **\"🔗 קישור מותאם אישית...\"**. "
        "יופיעו שני שדות:\n\n"
        "- **שם הפרויקט** - שם קבוע וייחודי לפרויקט הזה (למשל \"WISLAB\"). זה מה שמפריד "
        "בין המלאי/ה-WIP/תוכנית ההרכבה של הפרויקטים השונים - תשתמש **תמיד באותו שם** "
        "לפרויקט הזה, גם אם שם קובץ ה-Excel עצמו משתנה מדי חודש (למשל עם תאריך בשם).\n"
        "- **כתובת הקובץ** - הדבק כאן את הקישור מ-GitHub שהעתקת בשלב 1.\n\n"
        "לאחר שממלאים את שני השדות, יופיע כפתור **\"💾 שמור פרויקט זה לקבע\"** - לחיצה עליו "
        "שומרת את הפרויקט בענן, וממנה והלאה הוא יופיע ברשימה הקבועה בסיידבר בכל כניסה "
        "לאפליקציה, בלי לחזור על התהליך."
    )
    st.markdown(
        "**איך מוודאים שזה נטען נכון:** מיד אחרי הבחירה, מופיעה שורת אימות קטנה בסיידבר "
        "עם מספר הפריטים שזוהו בעץ, מספר החודשים, ושמות המוצרים ברמה 0 - תבדוק שזה תואם "
        "למה שאתה מצפה למצוא בקובץ שהעלית."
    )

elif nav_page == "🛒 דשבורד רכש חוצה-פרויקטים":
    st.markdown('<div class="section-title">🛒 דשבורד רכש חוצה-פרויקטים</div>', unsafe_allow_html=True)
    st.caption(
        "מציג את כל הביקוש החודשי (= מה שנרכוש בפועל, ללא נטרול מול מלאי/ETA) מכל הפרויקטים "
        "הזמינים בו-זמנית. מחיר היחידה נלקח בנפרד לכל פרויקט - אותו מק\"ט יכול להופיע במחירים "
        "שונים בפרויקטים שונים."
    )

    with st.spinner("טוען נתוני רכש מכל הפרויקטים... (זה יכול לקחת רגע בפעם הראשונה)"):
        procurement_df, failed_projects = load_all_projects_procurement(AVAILABLE_PROJECTS)

    if failed_projects:
        with st.expander(f"⚠️ {len(failed_projects)} פרויקטים לא נכללו בדשבורד - לחץ לפרטים"):
            for p_name, p_err in failed_projects:
                st.caption(f"**{p_name}**: {p_err}")

    if procurement_df.empty:
        st.info("לא נמצאו נתוני רכש בשום פרויקט זמין.")
    else:
        with st.container():
            st.markdown("#### 🔎 סינון")
            f_col1, f_col2, f_col3, f_col4 = st.columns(4)
            with f_col1:
                f_projects = st.multiselect(
                    "פרויקטים", sorted(procurement_df["Project"].unique()),
                    default=sorted(procurement_df["Project"].unique()), key="proc_f_projects"
                )
            with f_col2:
                f_suppliers = st.multiselect(
                    "ספקים (ריק = הכל)", sorted(procurement_df["Supplier"].unique()), key="proc_f_suppliers"
                )
            with f_col3:
                all_months_sorted = sorted(procurement_df["YearMonth"].unique())
                f_month_range = st.select_slider(
                    "טווח חודשים", options=all_months_sorted,
                    value=(all_months_sorted[0], all_months_sorted[-1]), key="proc_f_months"
                )
            with f_col4:
                f_search = st.text_input("חיפוש חופשי (מק\"ט / תיאור)", key="proc_f_search")

        filtered = procurement_df[procurement_df["Project"].isin(f_projects)] if f_projects else procurement_df.iloc[0:0]
        if f_suppliers:
            filtered = filtered[filtered["Supplier"].isin(f_suppliers)]
        filtered = filtered[(filtered["YearMonth"] >= f_month_range[0]) & (filtered["YearMonth"] <= f_month_range[1])]
        if f_search:
            s_lower = f_search.strip().lower()
            filtered = filtered[
                filtered["PN"].str.lower().str.contains(s_lower, na=False) |
                filtered["Description"].str.lower().str.contains(s_lower, na=False)
            ]

        if filtered.empty:
            st.warning("אין נתונים התואמים את הסינון הנוכחי.")
        else:
            total_value = filtered["Value"].sum()
            total_qty = filtered["Quantity"].sum()
            n_suppliers = filtered["Supplier"].nunique()
            n_pns = filtered["PN"].nunique()
            n_projects = filtered["Project"].nunique()

            kcol1, kcol2, kcol3, kcol4, kcol5 = st.columns(5)
            with kcol1:
                kpi_card("💰 שווי רכש כולל", f"${total_value:,.0f}", f"{len(filtered):,} שורות ביקוש", "green")
            with kcol2:
                kpi_card("📦 כמות כוללת", f"{total_qty:,.0f}", "יחידות", "blue")
            with kcol3:
                kpi_card("🏭 ספקים", n_suppliers, "ספקים ייחודיים", "orange")
            with kcol4:
                kpi_card("🔧 מק\"טים", n_pns, "פריטים ייחודיים", "blue")
            with kcol5:
                kpi_card("📁 פרויקטים", n_projects, "בטווח הנבחר", "green")

            st.markdown("#### 📈 מגמת רכש חודשית")
            monthly = filtered.groupby(["YearMonth", "Project"], as_index=False)["Value"].sum()
            fig_trend = px.bar(
                monthly, x="YearMonth", y="Value", color="Project",
                labels={"YearMonth": "חודש", "Value": "שווי רכש ($)", "Project": "פרויקט"},
                color_discrete_sequence=COLOR_SEQ
            )
            fig_trend.update_layout(template=PLOTLY_TEMPLATE, height=380, margin=dict(t=20, b=10, l=10, r=10), legend_title_text="")
            st.plotly_chart(fig_trend, use_container_width=True)

            chart_col1, chart_col2 = st.columns(2)
            with chart_col1:
                st.markdown("#### 🏭 שווי רכש לפי ספק (Top 15)")
                by_supplier = filtered.groupby("Supplier", as_index=False)["Value"].sum().sort_values("Value", ascending=False).head(15)
                fig_sup = px.bar(
                    by_supplier.sort_values("Value"), x="Value", y="Supplier", orientation="h",
                    color="Value", color_continuous_scale="Tealgrn",
                    labels={"Value": "שווי רכש ($)", "Supplier": ""}
                )
                fig_sup.update_layout(template=PLOTLY_TEMPLATE, height=420, margin=dict(t=10, b=10, l=10, r=10), coloraxis_showscale=False)
                st.plotly_chart(fig_sup, use_container_width=True)
            with chart_col2:
                st.markdown("#### 📁 שווי רכש לפי פרויקט")
                by_project = filtered.groupby("Project", as_index=False)["Value"].sum().sort_values("Value", ascending=False)
                fig_proj = px.pie(
                    by_project, names="Project", values="Value", hole=0.5,
                    color_discrete_sequence=COLOR_SEQ
                )
                fig_proj.update_layout(template=PLOTLY_TEMPLATE, height=420, margin=dict(t=10, b=10, l=10, r=10))
                st.plotly_chart(fig_proj, use_container_width=True)

            st.markdown("#### 🗓️ מפת חום: ספק מול חודש (שווי רכש)")
            top_suppliers_for_heatmap = filtered.groupby("Supplier")["Value"].sum().sort_values(ascending=False).head(20).index
            heatmap_src = filtered[filtered["Supplier"].isin(top_suppliers_for_heatmap)]
            if not heatmap_src.empty:
                pivot = heatmap_src.pivot_table(index="Supplier", columns="YearMonth", values="Value", aggfunc="sum", fill_value=0)
                pivot = pivot.reindex(top_suppliers_for_heatmap)
                fig_heat = px.imshow(
                    pivot, aspect="auto", color_continuous_scale="Blues",
                    labels={"x": "חודש", "y": "ספק", "color": "שווי ($)"}
                )
                fig_heat.update_layout(template=PLOTLY_TEMPLATE, height=max(400, 24 * len(pivot)), margin=dict(t=10, b=10, l=10, r=10))
                st.plotly_chart(fig_heat, use_container_width=True)

            st.markdown("#### 🔝 Top 15 פריטים לפי שווי")
            top_items = filtered.groupby(["PN", "Description"], as_index=False).agg(
                Total_Value=("Value", "sum"), Total_Qty=("Quantity", "sum")
            ).sort_values("Total_Value", ascending=False).head(15)
            fig_items = px.bar(
                top_items.sort_values("Total_Value"), x="Total_Value", y="PN", orientation="h",
                color="Total_Value", color_continuous_scale="Purples",
                hover_data=["Description", "Total_Qty"],
                labels={"Total_Value": "שווי רכש ($)", "PN": ""}
            )
            fig_items.update_layout(template=PLOTLY_TEMPLATE, height=420, margin=dict(t=10, b=10, l=10, r=10), coloraxis_showscale=False)
            st.plotly_chart(fig_items, use_container_width=True)

            st.markdown("#### 📋 טבלה מפורטת (ניתנת לסינון/מיון/הורדה)")
            display_cols = ["Project", "PN", "Description", "Supplier", "YearMonth", "Quantity", "Unit_Price", "Value"]
            display_df = filtered[display_cols].rename(columns={
                "Project": "פרויקט", "PN": "מק\"ט", "Description": "תיאור", "Supplier": "ספק",
                "YearMonth": "חודש", "Quantity": "כמות", "Unit_Price": "מחיר יחידה ($)", "Value": "שווי כולל ($)"
            }).sort_values("שווי כולל ($)", ascending=False)
            st.dataframe(display_df, use_container_width=True, height=420)
            st.download_button(
                "⬇️ הורד את הטבלה כ-CSV",
                data=display_df.to_csv(index=False).encode("utf-8-sig"),
                file_name="procurement_dashboard.csv",
                mime="text/csv"
            )
