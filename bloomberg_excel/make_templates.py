"""Build Bloomberg Excel templates (BDH/BDP formulas) covering the same data as bbg_pull.py.

Open each workbook on a Bloomberg Terminal PC with the Excel add-in; the formulas pull the data on open.
Every data block has the same layout so the filled workbooks can be read back in one way:
    row 1  ticker            row 3  "Date" + field mnemonics
    row 2  description/meta  row 4  the BDH formula (data fills down from here)

Run:  python bloomberg_excel/make_templates.py   (writes bloomberg_excel/templates/*.xlsx)
"""
import datetime as dt
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as L

OUT = Path(__file__).resolve().parent / "templates"
AS_OF = dt.date(2026, 10, 13)
MONTHS = "Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec".split()
CODES = "FGHJKMNQUVXZ"

FONT = Font(name="Arial", size=10)
BOLD = Font(name="Arial", size=10, bold=True)
TITLE = Font(name="Arial", size=14, bold=True)
INPUT = Font(name="Arial", size=10, color="0000FF")
FILL = PatternFill("solid", fgColor="FFFF00")
HEAD = PatternFill("solid", fgColor="D9E1F2")

DAILY_OPTS = '"Dir=V","Dts=S","Sort=A","Days=T","Fill=B"'
UNADJ_OPTS = '"CshAdjNormal=N","CshAdjAbnormal=N","CapChg=N"'

# ------------------------------------------------------------------------------------------------ data sets
YEN_CORE = [
    ("USDJPY Curncy", ["PX_LAST", "PX_HIGH", "PX_LOW"], "USD/JPY spot"),
    ("USDJPYV1W Curncy", ["PX_LAST"], "1-week ATM implied vol"),
    ("USDJPYV1M Curncy", ["PX_LAST", "PX_BID", "PX_ASK"], "1-month ATM implied vol"),
    ("USDJPYV3M Curncy", ["PX_LAST"], "3-month ATM implied vol"),
    ("USDJPY25R1M Curncy", ["PX_LAST"], "1-month 25-delta risk reversal"),
    ("USDJPY25B1M Curncy", ["PX_LAST"], "1-month 25-delta butterfly"),
    ("USDJPY10R1M Curncy", ["PX_LAST"], "1-month 10-delta risk reversal"),
    ("USDJPY10B1M Curncy", ["PX_LAST"], "1-month 10-delta butterfly"),
    ("JPY1M Curncy", ["PX_LAST"], "1-month forward points"),
]
YEN_EXTRA = [
    ("USDJPYV2W Curncy", ["PX_LAST"], "2-week ATM implied vol"),
    ("USDJPYV2M Curncy", ["PX_LAST"], "2-month ATM implied vol"),
    ("USDJPYV6M Curncy", ["PX_LAST"], "6-month ATM implied vol"),
    ("USDJPYV1Y Curncy", ["PX_LAST"], "1-year ATM implied vol"),
    ("USDJPY25R1W Curncy", ["PX_LAST"], "1-week 25-delta risk reversal"),
    ("USDJPY25R3M Curncy", ["PX_LAST"], "3-month 25-delta risk reversal"),
    ("USDJPY25B3M Curncy", ["PX_LAST"], "3-month 25-delta butterfly"),
    ("JPY1W Curncy", ["PX_LAST"], "1-week forward points"),
    ("JPY3M Curncy", ["PX_LAST"], "3-month forward points"),
    ("JY1 Curncy", ["PX_LAST", "PX_VOLUME", "OPEN_INT", "30DAY_IMPVOL_100.0%MNY_DF"], "CME yen futures, front"),
    ("JY2 Curncy", ["PX_LAST", "OPEN_INT"], "CME yen futures, second"),
    ("USGG2YR Index", ["PX_LAST"], "US 2-year Treasury yield"),
    ("USGG10YR Index", ["PX_LAST"], "US 10-year Treasury yield"),
    ("GTJPY2Y Govt", ["PX_LAST"], "Japan 2-year JGB yield"),
    ("GTJPY10Y Govt", ["PX_LAST"], "Japan 10-year JGB yield"),
    ("CL1 Comdty", ["PX_LAST"], "WTI crude, front month"),
    ("CO1 Comdty", ["PX_LAST"], "Brent crude, front month"),
    ("OVX Index", ["PX_LAST"], "Cboe crude oil volatility index"),
    ("VIX Index", ["PX_LAST"], "Cboe VIX"),
]
FUT_FIELDS = ["PX_OPEN", "PX_HIGH", "PX_LOW", "PX_LAST", "PX_SETTLE", "PX_VOLUME", "OPEN_INT"]
EQ_FIELDS = ["PX_OPEN", "PX_HIGH", "PX_LOW", "PX_LAST", "PX_VOLUME"]
OIL_GENERICS = (
    [(f"CL{i} Comdty", FUT_FIELDS, f"WTI crude, generic month {i}") for i in range(1, 13)]
    + [(f"CO{i} Comdty", FUT_FIELDS, f"Brent crude, generic month {i}") for i in range(1, 7)]
    + [("HO1 Comdty", FUT_FIELDS, "NY Harbor ULSD (heating oil), front"),
       ("XB1 Comdty", FUT_FIELDS, "RBOB gasoline, front"),
       ("QM1 Comdty", FUT_FIELDS, "E-mini WTI, front"),
       ("NG1 Comdty", FUT_FIELDS, "Henry Hub natural gas, front")])
OIL_EQUITIES = {
    "USO": "US Oil Fund (front WTI)", "BNO": "US Brent Oil Fund", "USL": "US 12-Month Oil Fund",
    "DBO": "Invesco DB Oil", "UCO": "ProShares 2x crude", "SCO": "ProShares -2x crude",
    "XLE": "Energy Select SPDR", "XOP": "S&P Oil & Gas E&P ETF", "OIH": "Oil services ETF",
    "VDE": "Vanguard energy ETF", "CRAK": "Oil refiners ETF", "XOM": "ExxonMobil", "CVX": "Chevron",
    "COP": "ConocoPhillips", "OXY": "Occidental", "EOG": "EOG Resources", "DVN": "Devon",
    "FANG": "Diamondback", "APA": "APA Corp", "DAL": "Delta Air Lines", "UAL": "United Airlines",
    "AAL": "American Airlines", "LUV": "Southwest", "JETS": "Airlines ETF", "SPY": "S&P 500 ETF",
}
MACRO = [("OVX Index", ["PX_OPEN", "PX_HIGH", "PX_LOW", "PX_LAST"], "Cboe crude oil volatility index"),
         ("VIX Index", ["PX_LAST"], "Cboe VIX"),
         ("SPX Index", ["PX_LAST"], "S&P 500"),
         ("USGG3M Index", ["PX_LAST"], "US 3-month T-bill yield (risk-free rate)"),
         ("USGG1YR Index", ["PX_LAST"], "US 1-year Treasury yield"),
         ("USGG2YR Index", ["PX_LAST"], "US 2-year Treasury yield"),
         ("DXY Curncy", ["PX_LAST"], "US dollar index")]
VOL_FIELDS = (
    [f"{t}_IMPVOL_{m}%MNY_DF" for t in ("30DAY", "60DAY", "3MTH", "6MTH")
     for m in ("90.0", "95.0", "100.0", "105.0", "110.0")]
    + [f"{t}_{cp}_IMP_VOL_{d}DELTA_DFLT" for t in ("1M", "2M", "3M") for cp in ("CALL", "PUT")
       for d in ("10", "25", "50")]
    + ["HIST_CALL_IMP_VOL", "HIST_PUT_IMP_VOL", "VOLATILITY_10D", "VOLATILITY_30D", "VOLATILITY_90D",
       "PUT_CALL_VOLUME_RATIO_CUR_DAY", "OPEN_INT_TOTAL_CALL", "OPEN_INT_TOTAL_PUT",
       "VOLUME_TOTAL_CALL", "VOLUME_TOTAL_PUT"])
VOL_GROUPS = {
    "A": ["CL1 Comdty", "CL2 Comdty", "CO1 Comdty", "USO US Equity"],
    "B": ["BNO US Equity", "XLE US Equity", "XOP US Equity", "OXY US Equity"],
    "C": ["XOM US Equity", "DAL US Equity", "JETS US Equity", "SPY US Equity"],
}
OPT_FIELDS = ["PX_BID", "PX_ASK", "PX_MID", "PX_LAST", "PX_SETTLE", "PX_VOLUME", "OPEN_INT", "IVOL_MID",
              "IVOL_BID", "IVOL_ASK", "DELTA_MID", "GAMMA_MID", "VEGA_MID", "THETA_MID", "OPT_UNDL_PX"]
OPT_EQUITIES = ["USO", "XLE", "XOP", "OXY", "DAL"]
INTRADAY = ["CL1 Comdty", "CO1 Comdty", "USO US Equity"]
INTRADAY_FIELDS = ["OPEN", "HIGH", "LOW", "LAST_PRICE", "VOLUME", "NUMBER_TICKS"]


# ------------------------------------------------------------------------------------------------ helpers
def new_book():
    wb = Workbook()
    wb.calculation.fullCalcOnLoad = True  # Excel computes every formula when the file opens
    return wb


def put(ws, ref, value, font=FONT, fill=None, fmt=None):
    c = ws[ref]
    c.value = value
    c.font = font
    if fill:
        c.fill = fill
    if fmt:
        c.number_format = fmt
    return c


def settings_sheet(wb, title, steps, inputs):
    """First sheet: title, instructions and the yellow input cells. inputs: [(label, value, fmt, note)]."""
    ws = wb.active
    ws.title = "Settings"
    put(ws, "A1", title, TITLE)
    row = 3
    for label, value, fmt, note in inputs:
        put(ws, f"A{row}", label, BOLD)
        put(ws, f"B{row}", value, INPUT, FILL, fmt)
        put(ws, f"C{row}", note)
        row += 1
    row += 1
    put(ws, f"A{row}", "Legend", BOLD)
    put(ws, f"A{row + 1}", "Yellow cells with blue text", INPUT, FILL)
    put(ws, f"B{row + 1}", "inputs you may change; everything else is a formula")
    row += 3
    put(ws, f"A{row}", "How to use", BOLD)
    for i, s in enumerate(steps, 1):
        put(ws, f"A{row + i}", f"{i}. {s}")
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 70
    return ws


COMMON_STEPS = [
    "Open this file in Excel on the Bloomberg PC with the Terminal logged in. Click Enable Editing / Enable Content.",
    "Wait until no cell says 'Requesting Data...'. If nothing loads: Bloomberg tab > Refresh > Refresh Workbook.",
    "A cell reading '#N/A Invalid Security' or '#N/A Field Not Applicable' just means that item is unavailable.",
    "File > Save As, keep .xlsx, and add the date to the name. Upload the saved file to Claude.",
]


def block(ws, col, ticker, fields, desc, start, end, opts=DAILY_OPTS, ticker_is_formula=False, meta=None):
    """One security: ticker / description / headers / BDH. Returns the next free column (one blank gap)."""
    c0 = L(col)
    put(ws, f"{c0}1", ticker, FONT if ticker_is_formula else INPUT, None if ticker_is_formula else FILL)
    put(ws, f"{c0}2", meta if meta else desc)
    if meta:
        put(ws, f"{L(col + 1)}2", desc)
    put(ws, f"{c0}3", "Date", BOLD, HEAD)
    for i, f in enumerate(fields, 1):
        put(ws, f"{L(col + i)}3", f, BOLD, HEAD)
    last = L(col + len(fields))
    put(ws, f"{c0}4", f"=BDH({c0}$1,{L(col + 1)}$3:{last}$3,{start},{end},{opts})")
    ws.column_dimensions[c0].width = 12
    for i in range(1, len(fields) + 1):
        ws.column_dimensions[L(col + i)].width = max(10, min(len(fields[i - 1]) + 2, 28))
    return col + len(fields) + 2


START = "Settings!$B$3"
END = 'IF(Settings!$B$4="","",Settings!$B$4)'
DATE_FMT = "yyyy-mm-dd"


def date_inputs(start, end_note="Leave blank for today"):
    return [("Start date", start, DATE_FMT, "First date of history"),
            ("End date", None, DATE_FMT, end_note)]


def series_sheet(wb, name, items, opts=DAILY_OPTS):
    ws = wb.create_sheet(name)
    col = 1
    for ticker, fields, desc in items:
        col = block(ws, col, ticker, fields, desc, START, END, opts)
    ws.freeze_panes = "A4"
    return ws


def save(wb, name):
    OUT.mkdir(parents=True, exist_ok=True)
    wb.save(OUT / name)
    print("wrote", name)


# ------------------------------------------------------------------------------------------------ workbooks
def test_book():
    wb = new_book()
    settings_sheet(wb, "Bloomberg Excel test (open this one first)", [
        "Open in Excel on the Bloomberg PC with the Terminal logged in.",
        "Go to the Test sheet. Cell B1 should show the WTI price and A4 down should fill with ~2 weeks of prices.",
        "If you see numbers, Excel works and you can open the other templates.",
        "If you see 'Workflow review needed' or #NAME?, Excel downloads are blocked too: contact HELP HELP.",
    ], date_inputs(AS_OF - dt.timedelta(14)))
    ws = wb.create_sheet("Test")
    put(ws, "A1", "Latest WTI price:", BOLD)
    put(ws, "B1", '=BDP("CL1 Comdty","PX_LAST")')
    block(ws, 4, "CL1 Comdty", ["PX_LAST"], "WTI crude, front month", START, END)
    save(wb, "00_Test.xlsx")


def yen_book():
    wb = new_book()
    settings_sheet(wb, "Yen: spot, FX-option vol surface, forwards, rates (yen-core + yen-all)",
                   COMMON_STEPS, date_inputs(dt.date(2010, 1, 1)))
    series_sheet(wb, "Yen Core", YEN_CORE)
    series_sheet(wb, "Yen Extra", YEN_EXTRA)
    save(wb, "01_Yen.xlsx")


def oil_core_book():
    wb = new_book()
    settings_sheet(wb, "Oil core: futures generics, oil ETFs and stocks, vol indices, rates (oil-core)",
                   COMMON_STEPS, date_inputs(dt.date(2010, 1, 1)))
    series_sheet(wb, "Futures", OIL_GENERICS)
    eqs = [(f"{t} US Equity", EQ_FIELDS, name) for t, name in OIL_EQUITIES.items()]
    series_sheet(wb, "Equities", eqs)
    series_sheet(wb, "Equities Unadjusted", [(t, ["PX_LAST"], d + " (unadjusted close)") for t, _, d in eqs],
                 opts=f"{DAILY_OPTS},{UNADJ_OPTS}")
    series_sheet(wb, "Macro", MACRO)
    save(wb, "02_Oil_Core.xlsx")


def oil_vol_books():
    for g, secs in VOL_GROUPS.items():
        wb = new_book()
        settings_sheet(wb, f"Oil implied-vol surfaces, part {g} of 3 (oil-vol)",
                       COMMON_STEPS + ["This is a large pull. Open parts A, B and C on different days if Bloomberg "
                                       "warns about data limits."],
                       date_inputs(dt.date(2016, 1, 1)))
        for s in secs:
            ws = wb.create_sheet(s.split()[0])
            half = len(VOL_FIELDS) // 2  # two BDH calls of 24 fields each
            col = block(ws, 1, s, VOL_FIELDS[:half], "implied-vol surface (moneyness)", START, END)
            block(ws, col, s, VOL_FIELDS[half:], "implied-vol surface (delta, realized, volume)", START, END)
            ws.freeze_panes = "B4"
        save(wb, f"03{g}_Oil_Vol.xlsx")


def ltd_formula(root, y, m):
    """Approximate last trading day (Excel formula) for a CL or CO contract of delivery year y, month m."""
    if root == "CL":  # 3 business days before the 25th of the month before delivery
        return f"WORKDAY(DATE({y},{m - 1},25),-3)"
    return f"WORKDAY(DATE({y},{m - 1},1),-1)"  # CO: last business day of the 2nd month before delivery


def fut_ticker(root, y, m, ltd_cell, asof="Settings!$B$5"):
    # Expired contracts use a 2-digit year (CLZ25), live ones a 1-digit year (CLZ6).
    return (f'="{root}{CODES[m - 1]}"&IF({ltd_cell}<{asof},"{y % 100:02d}","{y % 10}")&" Comdty"')


def futures_book(root, name):
    wb = new_book()
    settings_sheet(wb, f"{name}: every individual futures contract since 2016 (oil-futures)",
                   COMMON_STEPS + [
                       "Row 1 ticker: expired contracts use a 2-digit year (CLZ25), live ones 1-digit (CLZ6). "
                       "It switches on the As-of date.",
                       "If one contract shows '#N/A Invalid Security', change As-of date to today and reopen.",
                   ],
                   date_inputs(dt.date(2016, 1, 1))
                   + [("As-of date", AS_OF, DATE_FMT, "Set to today's date; decides expired vs live tickers")])
    for y in range(2016, 2030):
        ws = wb.create_sheet(f"{root} {y}")
        col = 1
        for m in range(1, 13):
            # row 2: approximate last trading day (used for the ticker spelling) + Bloomberg's real one
            c0, c1 = L(col), L(col + 1)
            ticker = fut_ticker(root, y, m, f"{c0}2")
            col_next = block(ws, col, ticker, FUT_FIELDS, "", START, END, ticker_is_formula=True,
                             meta=f"={ltd_formula(root, y, m)}")
            ws[f"{c0}2"].number_format = DATE_FMT
            put(ws, f"{c1}2", f'=BDP({c0}$1,"LAST_TRADEABLE_DT")', fmt=DATE_FMT)
            col = col_next
        ws.freeze_panes = "A4"
    save(wb, f"04_{root}_Futures.xlsx" if root == "CL" else f"05_{root}_Futures.xlsx")


def option_inputs(ws, rows):
    """Per-sheet inputs/derived values in columns A:B rows 1.. ; data blocks start at column D."""
    for r, (label, value, fmt, is_input) in enumerate(rows, 1):
        put(ws, f"A{r}", label, BOLD)
        put(ws, f"B{r}", value, INPUT if is_input else FONT, FILL if is_input else None, fmt)
    ws.column_dimensions["A"].width = 22
    ws.column_dimensions["B"].width = 22


def cl_options_book():
    wb = new_book()
    settings_sheet(wb, "CL (WTI) monthly options: ATM +/- 8 strikes, last 100 days before expiry (oil-options)",
                   COMMON_STEPS[:3] + [
                       "Each sheet is one futures contract month of the Year below. Strikes are centred on the "
                       "futures price 45 days before expiry.",
                       "Save As CL_Options_<year>.xlsx, then change Year, wait for it to reload, and Save As again. "
                       "Do this for each year 2018-2026.",
                       "If every option shows '#N/A Invalid Security', change Strike format to 2 (e.g. 'CLZ5C 70.00') "
                       "and check one ticker on the Terminal: CLA Comdty OMON <GO>.",
                   ],
                   [("Year", 2026, "0", "Contract year (delivery month) to pull"),
                    ("Strike format", 1, "0", "1 = '70', 2 = '70.00'"),
                    ("As-of date", AS_OF, DATE_FMT, "Set to today's date; decides expired vs live tickers")])
    for m in range(1, 13):
        ws = wb.create_sheet(MONTHS[m - 1])
        y = "Settings!$B$3"
        option_inputs(ws, [
            ("Contract month", m, "0", False),
            ("Futures last trade", f"=WORKDAY(DATE({y},B1-1,25),-3)", DATE_FMT, False),
            ("Option expiry (approx)", "=WORKDAY(B2,-1)", DATE_FMT, False),
            ("Year in ticker", f'=IF(B2<Settings!$B$5,TEXT(MOD({y},100),"00"),TEXT(MOD({y},10),"0"))', None, False),
            ("Futures ticker", f'="CL"&MID("{CODES}",B1,1)&B4&" Comdty"', None, False),
            ("Strike reference date", "=B3-45", DATE_FMT, False),
            ("Futures price on that date", '=BDH(B5,"PX_LAST",B6,B6,"Days=A","Fill=P","Dts=H")', "0.00", False),
            ("ATM strike", "=ROUND(B7,0)", "0", False),
            ("History start", "=B3-100", DATE_FMT, False),
            ("History end", "=B3", DATE_FMT, False),
        ])
        col = 4
        for k in range(-8, 9):
            for cp in "CP":
                strike = f"(B$8{k:+d})" if k else "B$8"
                kfmt = f'IF(Settings!$B$4=2,TEXT({strike},"0.00"),{strike})'
                ticker = f'="CL"&MID("{CODES}",B$1,1)&B$4&"{cp} "&{kfmt}&" Comdty"'
                col = block(ws, col, ticker, OPT_FIELDS, f"{'call' if cp == 'C' else 'put'} {k:+d} from ATM",
                            "$B$9", "$B$10", ticker_is_formula=True)
        ws.freeze_panes = "D4"
    save(wb, "06_CL_Options.xlsx")


def equity_options_book(und):
    wb = new_book()
    settings_sheet(wb, f"{und} monthly options: ATM +/- 5 strikes, last 100 days before expiry (oil-equity-options)",
                   COMMON_STEPS[:3] + [
                       "Each sheet is one monthly expiry (3rd Friday) of the Year below. Strikes are centred on the "
                       "unadjusted close 45 days before expiry.",
                       f"Save As {und}_Options_<year>.xlsx, then change Year, wait for it to reload, and Save As "
                       "again. Do this for each year 2018-2026.",
                       "If a whole sheet shows '#N/A Invalid Security', that Friday was a holiday: type the Thursday "
                       "date over the Expiry cell (B3) on that sheet.",
                   ],
                   [("Year", 2026, "0", "Expiry year to pull"),
                    ("Underlying", und, None, "US ticker")])
    for m in range(1, 13):
        ws = wb.create_sheet(MONTHS[m - 1])
        y = "Settings!$B$3"
        option_inputs(ws, [
            ("Month", m, "0", False),
            ("Underlying", '=Settings!$B$4&" US Equity"', None, False),
            ("Expiry (3rd Friday)", f"=DATE({y},B1,15)+MOD(6-WEEKDAY(DATE({y},B1,15)),7)", DATE_FMT, True),
            ("Strike reference date", "=B3-45", DATE_FMT, False),
            ("Unadjusted close then",
             f'=BDH(B2,"PX_LAST",B4,B4,"Days=A","Fill=P","Dts=H",{UNADJ_OPTS})', "0.00", False),
            ("Strike step", "=IF(B5<25,0.5,IF(B5<250,1,5))", "0.0", False),
            ("ATM strike", "=ROUND(B5/B6,0)*B6", "0.0", False),
            ("History start", "=B3-100", DATE_FMT, False),
            ("History end", "=B3", DATE_FMT, False),
        ])
        col = 4
        for k in range(-5, 6):
            for cp in "CP":
                strike = f"(B$7{k:+d}*B$6)" if k else "B$7"
                ticker = f'=Settings!$B$4&" US "&TEXT(B$3,"mm/dd/yy")&" {cp}"&{strike}&" Equity"'
                col = block(ws, col, ticker, OPT_FIELDS, f"{'call' if cp == 'C' else 'put'} {k:+d} from ATM",
                            "$B$8", "$B$9", ticker_is_formula=True)
        ws.freeze_panes = "D4"
    save(wb, f"07_{und}_Options.xlsx")


def intraday_book():
    wb = new_book()
    settings_sheet(wb, "5-minute bars for CL1, CO1, USO (oil-intraday)",
                   COMMON_STEPS + ["Bloomberg keeps only about 140 days of intraday bars, so older starts return "
                                   "nothing."],
                   [("Start date", AS_OF - dt.timedelta(140), DATE_FMT, "About 140 days ago is the limit"),
                    ("End date", AS_OF, DATE_FMT, "Set to today"),
                    ("Bar size (minutes)", 5, "0", "")])
    ws = wb.create_sheet("Intraday")
    col = 1
    opts = '"BarTp=T","BarSz="&Settings!$B$5,"Dir=V","Dts=S","Sort=A"'
    for s in INTRADAY:
        col = block(ws, col, s, INTRADAY_FIELDS, "5-minute trade bars", START, "Settings!$B$4", opts)
        ws.column_dimensions[L(col - len(INTRADAY_FIELDS) - 2)].width = 17
    ws.freeze_panes = "A4"
    save(wb, "08_Intraday.xlsx")


if __name__ == "__main__":
    test_book()
    yen_book()
    oil_core_book()
    oil_vol_books()
    futures_book("CL", "WTI crude (CL)")
    futures_book("CO", "Brent crude (CO)")
    cl_options_book()
    for u in OPT_EQUITIES:
        equity_options_book(u)
    intraday_book()
