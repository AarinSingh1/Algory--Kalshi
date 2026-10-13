"""Build Bloomberg Excel templates (BDH/BDP formulas) covering the same data as bbg_pull.py.

One workbook, one sheet per data set, and a Settings sheet with a Load switch per section so only the sections
switched on request data. Open it on a Bloomberg Terminal PC with the Excel add-in. Every data block has the same layout so the filled workbooks can be read back in one way:
    row 1  ticker            row 3  "Date" + field mnemonics
    row 2  description/meta  row 4  the BDH formula (data fills down from here)

Run:  python bloomberg_excel/make_templates.py   (writes bloomberg_excel/templates/Bloomberg_Data_Pull.xlsx)
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
def put(ws, ref, value, font=FONT, fill=None, fmt=None):
    c = ws[ref]
    c.value = value
    c.font = font
    if fill:
        c.fill = fill
    if fmt:
        c.number_format = fmt
    return c


DATE_FMT = "yyyy-mm-dd"
ASOF = "Settings!$B$3"
TABLE_ROW = 6  # first section row of the Settings table


class Section:
    """One row of the Settings table: Load switch (B), start (C), end (D), year (E), other (F)."""

    def __init__(self, row):
        self.load = f'Settings!$B${row}="Y"'
        self.start = f"Settings!$C${row}"
        self.end = f'IF(Settings!$D${row}="","",Settings!$D${row})'
        self.end_raw = f"Settings!$D${row}"
        self.year = f"Settings!$E${row}"
        self.other = f"Settings!$F${row}"

    def gate(self, expr):
        """A ticker that is blank unless this section's Load switch is Y, so switched-off sections request nothing."""
        return f'IF({self.load},{expr},"")'


def block(ws, col, ticker_expr, fields, desc, start, end, opts=DAILY_OPTS, row2=None):
    """One security: ticker / description / headers / BDH. Returns the next free column (one blank gap)."""
    c0 = L(col)
    put(ws, f"{c0}1", "=" + ticker_expr, BOLD)
    put(ws, f"{c0}2", row2 if row2 else desc, fmt=DATE_FMT if row2 else None)
    if row2:
        put(ws, f"{L(col + 1)}2", desc)
    put(ws, f"{c0}3", "Date", BOLD, HEAD)
    for i, f in enumerate(fields, 1):
        put(ws, f"{L(col + i)}3", f, BOLD, HEAD)
    put(ws, f"{c0}4", f"=BDH({c0}$1,{L(col + 1)}$3:{L(col + len(fields))}$3,{start},{end},{opts})")
    ws.column_dimensions[c0].width = 12
    for i, f in enumerate(fields, 1):
        ws.column_dimensions[L(col + i)].width = max(10, min(len(f) + 2, 28))
    return col + len(fields) + 2


def series_sheet(wb, name, sec, items, opts=DAILY_OPTS):
    ws = wb.create_sheet(name)
    col = 1
    for ticker, fields, desc in items:
        col = block(ws, col, sec.gate(f'"{ticker}"'), fields, desc, sec.start, sec.end, opts)
    ws.freeze_panes = "A4"
    return ws


def side_inputs(ws, rows):
    """Derived values for option sheets in A:B; data blocks start at column D."""
    for r, (label, value, fmt) in enumerate(rows, 1):
        put(ws, f"A{r}", label, BOLD)
        put(ws, f"B{r}", value, fmt=fmt)
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 22


def ltd_formula(root, y, m):
    """Approximate last trading day (Excel formula) for a CL or CO contract of delivery year y, month m."""
    if root == "CL":  # 3 business days before the 25th of the month before delivery
        return f"WORKDAY(DATE({y},{m}-1,25),-3)"
    return f"WORKDAY(DATE({y},{m}-1,1),-1)"  # CO: last business day of the 2nd month before delivery


# ------------------------------------------------------------------------------------------------ sections
def sec_test(wb, sec):
    ws = wb.create_sheet("Test")
    put(ws, "A1", "Latest WTI price:", BOLD)
    put(ws, "B1", f'=IF({sec.load},BDP("CL1 Comdty","PX_LAST"),"")')
    block(ws, 4, sec.gate('"CL1 Comdty"'), ["PX_LAST"], "WTI crude, front month", sec.start, sec.end)


def sec_yen(wb, sec):
    series_sheet(wb, "Yen Core", sec, YEN_CORE)
    series_sheet(wb, "Yen Extra", sec, YEN_EXTRA)


def sec_oil_core(wb, sec):
    series_sheet(wb, "Oil Futures", sec, OIL_GENERICS)
    eqs = [(f"{t} US Equity", EQ_FIELDS, name) for t, name in OIL_EQUITIES.items()]
    series_sheet(wb, "Oil Equities", sec, eqs)
    series_sheet(wb, "Oil Equities Unadj", sec, [(t, ["PX_LAST"], d + " (unadjusted close)") for t, _, d in eqs],
                 opts=f"{DAILY_OPTS},{UNADJ_OPTS}")
    series_sheet(wb, "Macro", sec, MACRO)


def sec_vol(group):
    def build(wb, sec):
        half = len(VOL_FIELDS) // 2  # two BDH calls of 24 fields each
        for s in VOL_GROUPS[group]:
            ws = wb.create_sheet("Vol " + s.split()[0])
            t = sec.gate(f'"{s}"')
            col = block(ws, 1, t, VOL_FIELDS[:half], "implied-vol surface (moneyness)", sec.start, sec.end)
            block(ws, col, t, VOL_FIELDS[half:], "implied-vol surface (delta, realized, volume)", sec.start, sec.end)
            ws.freeze_panes = "B4"
    return build


def sec_futures(root):
    def build(wb, sec):
        for y in range(2016, 2030):
            ws = wb.create_sheet(f"{root} Fut {y}")
            col = 1
            for m in range(1, 13):
                c0, c1 = L(col), L(col + 1)
                # Expired contracts use a 2-digit year (CLZ25), live ones a 1-digit year (CLZ6); row 2 holds the
                # approximate last trading day that decides which, and C2 Bloomberg's actual one.
                ticker = sec.gate(f'"{root}{CODES[m - 1]}"&IF({c0}$2<{ASOF},"{y % 100:02d}","{y % 10}")&" Comdty"')
                nxt = block(ws, col, ticker, FUT_FIELDS, "", sec.start, sec.end, row2=f"={ltd_formula(root, y, m)}")
                put(ws, f"{c1}2", f'=IF({c0}$1="","",BDP({c0}$1,"LAST_TRADEABLE_DT"))', fmt=DATE_FMT)
                col = nxt
            ws.freeze_panes = "A4"
    return build


def sec_cl_options(wb, sec):
    y = sec.year
    for m in range(1, 13):
        ws = wb.create_sheet(f"CL Opt {MONTHS[m - 1]}")
        side_inputs(ws, [
            ("Contract month", m, "0"),
            ("Futures last trade", f"=WORKDAY(DATE({y},B1-1,25),-3)", DATE_FMT),
            ("Option expiry (approx)", "=WORKDAY(B2,-1)", DATE_FMT),
            ("Year in ticker", f'=IF(B2<{ASOF},TEXT(MOD({y},100),"00"),TEXT(MOD({y},10),"0"))', None),
            ("Futures ticker", sec.gate(f'"CL"&MID("{CODES}",B1,1)&B4&" Comdty"').join(["=", ""]), None),
            ("Strike reference date", "=B3-45", DATE_FMT),
            ("Futures price on that date",
             '=IF(B5="","",BDH(B5,"PX_LAST",B6,B6,"Days=A","Fill=P","Dts=H"))', "0.00"),
            ("ATM strike", '=IF(ISNUMBER(B7),ROUND(B7,0),"")', "0"),
            ("History start", "=B3-100", DATE_FMT),
            ("History end", "=B3", DATE_FMT),
        ])
        col = 4
        for k in range(-8, 9):
            for cp in "CP":
                strike = f"(B$8{k:+d})" if k else "B$8"
                kfmt = f'IF({sec.other}=2,TEXT({strike},"0.00"),{strike})'
                ticker = f'IF(ISNUMBER(B$8),"CL"&MID("{CODES}",B$1,1)&B$4&"{cp} "&{kfmt}&" Comdty","")'
                col = block(ws, col, ticker, OPT_FIELDS, f"{'call' if cp == 'C' else 'put'} {k:+d} from ATM",
                            "$B$9", "$B$10")
        ws.freeze_panes = "D4"


def sec_equity_options(und):
    def build(wb, sec):
        y = sec.year
        for m in range(1, 13):
            ws = wb.create_sheet(f"{und} Opt {MONTHS[m - 1]}")
            side_inputs(ws, [
                ("Month", m, "0"),
                ("Underlying", "=" + sec.gate(f'"{und} US Equity"'), None),
                ("Expiry (3rd Friday)", f"=DATE({y},B1,15)+MOD(6-WEEKDAY(DATE({y},B1,15)),7)", DATE_FMT),
                ("Strike reference date", "=B3-45", DATE_FMT),
                ("Unadjusted close then",
                 f'=IF(B2="","",BDH(B2,"PX_LAST",B4,B4,"Days=A","Fill=P","Dts=H",{UNADJ_OPTS}))', "0.00"),
                ("Strike step", '=IF(ISNUMBER(B5),IF(B5<25,0.5,IF(B5<250,1,5)),"")', "0.0"),
                ("ATM strike", '=IF(ISNUMBER(B5),ROUND(B5/B6,0)*B6,"")', "0.0"),
                ("History start", "=B3-100", DATE_FMT),
                ("History end", "=B3", DATE_FMT),
            ])
            ws["B3"].font, ws["B3"].fill = INPUT, FILL  # overwrite with the Thursday date if Friday was a holiday
            col = 4
            for k in range(-5, 6):
                for cp in "CP":
                    strike = f"(B$7{k:+d}*B$6)" if k else "B$7"
                    ticker = f'IF(ISNUMBER(B$7),"{und} US "&TEXT(B$3,"mm/dd/yy")&" {cp}"&{strike}&" Equity","")'
                    col = block(ws, col, ticker, OPT_FIELDS, f"{'call' if cp == 'C' else 'put'} {k:+d} from ATM",
                                "$B$8", "$B$9")
            ws.freeze_panes = "D4"
    return build


def sec_intraday(wb, sec):
    ws = wb.create_sheet("Intraday")
    col = 1
    opts = f'"BarTp=T","BarSz="&{sec.other},"Dir=V","Dts=S","Sort=A"'
    for s in INTRADAY:
        start_col = col
        col = block(ws, col, sec.gate(f'"{s}"'), INTRADAY_FIELDS, "5-minute trade bars", sec.start, sec.end_raw, opts)
        ws.column_dimensions[L(start_col)].width = 17
    ws.freeze_panes = "A4"


# (name, builder, start, end, year, other, note)
SECTIONS = [
    ("Test", sec_test, AS_OF - dt.timedelta(14), None, None, None,
     "Switch on first. Test sheet: B1 shows a price and column D fills with ~2 weeks of prices"),
    ("Yen", sec_yen, dt.date(2010, 1, 1), None, None, None, "Sheets: Yen Core, Yen Extra"),
    ("Oil core", sec_oil_core, dt.date(2010, 1, 1), None, None, None,
     "Sheets: Oil Futures, Oil Equities, Oil Equities Unadj, Macro"),
    ("Oil vol A", sec_vol("A"), dt.date(2016, 1, 1), None, None, None, "Vol sheets for CL1, CL2, CO1, USO (large)"),
    ("Oil vol B", sec_vol("B"), dt.date(2016, 1, 1), None, None, None, "Vol sheets for BNO, XLE, XOP, OXY (large)"),
    ("Oil vol C", sec_vol("C"), dt.date(2016, 1, 1), None, None, None, "Vol sheets for XOM, DAL, JETS, SPY (large)"),
    ("CL futures", sec_futures("CL"), dt.date(2016, 1, 1), None, None, None,
     "Every WTI contract 2016-2029, sheets CL Fut <year> (large)"),
    ("CO futures", sec_futures("CO"), dt.date(2016, 1, 1), None, None, None,
     "Every Brent contract 2016-2029, sheets CO Fut <year> (large)"),
    ("CL options", sec_cl_options, None, None, 2026, 1,
     "Year = contract year. Other = strike format: 1 gives 'CLZ6C 70', 2 gives 'CLZ6C 70.00'"),
] + [(f"{u} options", sec_equity_options(u), None, None, 2026, None, "Year = expiry year") for u in OPT_EQUITIES] + [
    ("Intraday", sec_intraday, AS_OF - dt.timedelta(140), AS_OF, None, 5,
     "Other = bar size in minutes. Bloomberg keeps ~140 days, so keep Start within that"),
]

STEPS = [
    "Open this file in Excel on the Bloomberg PC with the Terminal logged in. Click Enable Editing / Enable Content.",
    "Only sections with Load = Y request data. Test starts as Y; check the Test sheet shows numbers.",
    "Change Test to N and one other section to Y. Wait until no cell says 'Requesting Data...'.",
    "File > Save As .xlsx (save often). Then switch that section to N and the next to Y. Turning a section off "
    "clears its data, so Save As a new name per section, or keep each finished section on.",
    "Options sections cover one Year at a time: set the Year, let it load, Save As, then change the Year "
    "(2018-2026).",
    "'#N/A Invalid Security' or '#N/A Field Not Applicable' means that item isn't available; '#N/A Review' means "
    "Bloomberg is blocking downloads for this login (contact HELP HELP).",
    "If nothing loads: Bloomberg tab > Refresh > Refresh Workbook. Upload saved files to Claude.",
]


def build():
    wb = Workbook()
    wb.calculation.fullCalcOnLoad = True  # Excel computes every formula when the file opens
    ws = wb.active
    ws.title = "Settings"
    put(ws, "A1", "Bloomberg data pull: yen and oil (same data as bbg_pull.py)", TITLE)
    put(ws, "A3", "As-of date", BOLD)
    put(ws, "B3", AS_OF, INPUT, FILL, DATE_FMT)
    put(ws, "C3", "Set to today's date; decides expired vs live futures tickers")
    for i, h in enumerate(["Section", "Load (Y/N)", "Start date", "End date", "Year", "Other", "Notes"]):
        put(ws, f"{L(i + 1)}{TABLE_ROW - 1}", h, BOLD, HEAD)
    row = TABLE_ROW
    for name, fn, start, end, year, other, note in SECTIONS:
        put(ws, f"A{row}", name, BOLD)
        put(ws, f"B{row}", "Y" if name == "Test" else "N", INPUT, FILL)
        for c, v, fmt in (("C", start, DATE_FMT), ("D", end, DATE_FMT), ("E", year, "0"), ("F", other, "0")):
            put(ws, f"{c}{row}", v, INPUT, FILL, fmt)
        put(ws, f"G{row}", note)
        fn(wb, Section(row))
        row += 1
    put(ws, f"A{row}", "Blank End date = today.")
    row += 2
    put(ws, f"A{row}", "Legend", BOLD)
    put(ws, f"A{row + 1}", "Yellow cells, blue text", INPUT, FILL)
    put(ws, f"B{row + 1}", "inputs you may change; everything else is a formula")
    row += 3
    put(ws, f"A{row}", "How to use", BOLD)
    for i, s in enumerate(STEPS, 1):
        put(ws, f"A{row + i}", f"{i}. {s}")
    for c, w in zip("ABCDEFG", (16, 11, 12, 12, 7, 7, 80)):
        ws.column_dimensions[c].width = w
    ws.freeze_panes = f"A{TABLE_ROW}"
    OUT.mkdir(parents=True, exist_ok=True)
    wb.save(OUT / "Bloomberg_Data_Pull.xlsx")
    print(f"wrote Bloomberg_Data_Pull.xlsx ({len(wb.sheetnames)} sheets)")


if __name__ == "__main__":
    build()
