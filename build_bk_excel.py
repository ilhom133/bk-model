import zipfile, io, os, re

# ── helpers ──────────────────────────────────────────────────────────────────
def esc(s):
    return str(s).replace("&","&amp;").replace("<","&lt;").replace(">","&gt;").replace('"',"&quot;")

# Shared-string table  (we build it lazily)
sst_list = []
sst_map  = {}

def si(text):
    """Return 0-based index into the shared-string table."""
    text = str(text)
    if text not in sst_map:
        sst_map[text] = len(sst_list)
        sst_list.append(text)
    return sst_map[text]

# ── column-letter helper ──────────────────────────────────────────────────────
def col_letter(n):          # n is 1-based
    s = ""
    while n:
        n, r = divmod(n-1, 26)
        s = chr(65+r) + s
    return s

def cell_ref(row, col):     # both 1-based
    return f"{col_letter(col)}{row}"

# ── STYLE INDICES (we define all styles upfront) ──────────────────────────────
# xf indices used throughout:
#   0  normal
#   1  title  (white bold 16 on dark-blue bg)
#   2  section-header (white bold 11 on medium-blue)
#   3  col-header (white bold 10 on slate)
#   4  label  (bold 10)
#   5  number (number-format $#,##0 ; right-aligned)
#   6  total-row (bold, top+bottom border, number-format)
#   7  total-label (bold, top+bottom border)
#   8  note   (italic 9 grey)
#   9  ratio-ok  (green fill)
#  10  ratio-warn (amber fill)
#  11  ratio-bad  (red fill)
#  12  number plain right
#  13  check-green (bold green font)
#  14  section divider (light-blue fill, bold)
#  15  subtitle (italic 11 on dark-blue, white)

STYLES_XML = """\
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="3">
    <numFmt numFmtId="164" formatCode="&quot;$&quot;#,##0"/>
    <numFmt numFmtId="165" formatCode="&quot;$&quot;#,##0.00"/>
    <numFmt numFmtId="166" formatCode="0.00&quot;x&quot;"/>
  </numFmts>
  <fonts count="10">
    <font><sz val="10"/><name val="Calibri"/></font>
    <font><b/><sz val="16"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
    <font><b/><sz val="10"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
    <font><b/><sz val="10"/><name val="Calibri"/></font>
    <font><sz val="9"/><color rgb="FF595959"/><i/><name val="Calibri"/></font>
    <font><sz val="10"/><color rgb="FF1F6B0E"/><b/><name val="Calibri"/></font>
    <font><b/><sz val="10"/><color rgb="FF8B0000"/><name val="Calibri"/></font>
    <font><i/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>
    <font><b/><sz val="11"/><name val="Calibri"/></font>
  </fonts>
  <fills count="9">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1B3A6B"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF2E6DA4"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF4472C4"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFE2EFDA"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFF2CC"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFCE4D6"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFDAE3F3"/></patternFill></fill>
  </fills>
  <borders count="4">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border><left style="thin"><color auto="1"/></left><right style="thin"><color auto="1"/></right><top style="thin"><color auto="1"/></top><bottom style="thin"><color auto="1"/></bottom><diagonal/></border>
    <border><left/><right/><top style="medium"><color rgb="FF1B3A6B"/></top><bottom style="medium"><color rgb="FF1B3A6B"/></bottom><diagonal/></border>
    <border><left/><right/><top style="thin"><color auto="1"/></top><bottom style="thin"><color auto="1"/></bottom><diagonal/></border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="16">
    <xf numFmtId="0"   fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0"   fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="0"   fontId="2" fillId="3" borderId="0" xfId="0" applyFont="1" applyFill="1"><alignment horizontal="left" vertical="center"/></xf>
    <xf numFmtId="0"   fontId="3" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0"   fontId="4" fillId="0" borderId="3" xfId="0" applyFont="1" applyBorder="1"><alignment horizontal="left"/></xf>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="3" xfId="0" applyNumberFormat="1" applyBorder="1"><alignment horizontal="right"/></xf>
    <xf numFmtId="164" fontId="4" fillId="0" borderId="2" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1"><alignment horizontal="right"/></xf>
    <xf numFmtId="0"   fontId="4" fillId="0" borderId="2" xfId="0" applyFont="1" applyBorder="1"><alignment horizontal="left"/></xf>
    <xf numFmtId="0"   fontId="5" fillId="0" borderId="0" xfId="0" applyFont="1"><alignment horizontal="left" wrapText="1"/></xf>
    <xf numFmtId="0"   fontId="0" fillId="5" borderId="1" xfId="0" applyFill="1" applyBorder="1"><alignment horizontal="center"/></xf>
    <xf numFmtId="0"   fontId="0" fillId="6" borderId="1" xfId="0" applyFill="1" applyBorder="1"><alignment horizontal="center"/></xf>
    <xf numFmtId="0"   fontId="7" fillId="7" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"><alignment horizontal="center"/></xf>
    <xf numFmtId="0"   fontId="0" fillId="0" borderId="0" xfId="0"><alignment horizontal="right"/></xf>
    <xf numFmtId="0"   fontId="6" fillId="0" borderId="0" xfId="0" applyFont="1"><alignment horizontal="center"/></xf>
    <xf numFmtId="0"   fontId="4" fillId="8" borderId="0" xfId="0" applyFont="1" applyFill="1"><alignment horizontal="center"/></xf>
    <xf numFmtId="0"   fontId="8" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"><alignment horizontal="center" vertical="center"/></xf>
  </cellXfs>
</styleSheet>"""


# ── worksheet builder ─────────────────────────────────────────────────────────
class Sheet:
    def __init__(self, name):
        self.name   = name
        self.cells  = {}   # (row,col) -> (value, xf, is_string)
        self.merges = []   # "A1:D1"
        self.col_widths = {}  # col (1-based) -> width
        self.row_heights = {} # row (1-based) -> height

    def write(self, row, col, value, xf=0, is_string=True):
        self.cells[(row, col)] = (value, xf, is_string)

    def write_num(self, row, col, value, xf=5):
        self.cells[(row, col)] = (value, xf, False)

    def merge(self, r1, c1, r2, c2):
        self.merges.append(f"{cell_ref(r1,c1)}:{cell_ref(r2,c2)}")

    def col_w(self, col, w):
        self.col_widths[col] = w

    def row_h(self, row, h):
        self.row_heights[row] = h

    def to_xml(self):
        rows = {}
        for (r, c), (val, xf, is_str) in self.cells.items():
            rows.setdefault(r, []).append((c, val, xf, is_str))

        col_xml = ""
        for col, w in sorted(self.col_widths.items()):
            col_xml += f'<col min="{col}" max="{col}" width="{w}" customWidth="1"/>'

        rows_xml = ""
        for r in sorted(rows):
            ht = self.row_heights.get(r, "")
            ht_attr = f' ht="{ht}" customHeight="1"' if ht else ""
            rows_xml += f"<row r=\"{r}\"{ht_attr}>"
            for (c, val, xf, is_str) in sorted(rows[r]):
                ref = cell_ref(r, c)
                if is_str:
                    idx = si(val)
                    rows_xml += f'<c r="{ref}" t="s" s="{xf}"><v>{idx}</v></c>'
                else:
                    rows_xml += f'<c r="{ref}" s="{xf}"><v>{esc(val)}</v></c>'
            rows_xml += "</row>"

        merge_xml = ""
        if self.merges:
            merge_xml = "<mergeCells>" + "".join(f'<mergeCell ref="{m}"/>' for m in self.merges) + "</mergeCells>"

        return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
           xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetViews><sheetView workbookViewId="0" showGridLines="1"/></sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <cols>{col_xml}</cols>
  <sheetData>{rows_xml}</sheetData>
  {merge_xml}
</worksheet>"""

# ─────────────────────────────────────────────────────────────────────────────
# SHEET 1 — BALANCE SHEET
# ─────────────────────────────────────────────────────────────────────────────
def build_balance_sheet():
    s = Sheet("Balance Sheet")
    s.col_w(1, 5);  s.col_w(2, 38); s.col_w(3, 30); s.col_w(4, 16); s.col_w(5, 4)

    r = 1
    # Title banner
    s.row_h(r, 40)
    s.write(r, 1, "BURGER KING (Restaurant Brands International)", 1)
    s.merge(r, 1, r, 4)
    r += 1
    s.row_h(r, 22)
    s.write(r, 1, "Consolidated Balance Sheet  |  FY2025E  |  USD in Millions", 15)
    s.merge(r, 1, r, 4)
    r += 1
    s.write(r, 1, "Prepared under IFRS 16 / US GAAP  |  For Analytical Purposes Only", 8)
    s.merge(r, 1, r, 4)
    r += 2  # blank

    # Column headers
    s.row_h(r, 28)
    for c, h in [(2,"Line Item"),(3,"Notes / Basis"),(4,"USD (M)")]:
        s.write(r, c, h, 3)
    r += 1

    def section(title, row):
        s.row_h(row, 20)
        s.write(row, 1, "", 2); s.write(row, 2, title, 2); s.write(row, 3, "", 2); s.write(row, 4, "", 2)
        s.merge(row, 1, row, 4)

    def data_row(row, label, note, value, bold=False):
        lbl_xf = 4 if bold else 0
        s.write(row, 2, label, lbl_xf)
        s.write(row, 3, note,  8)
        if value != "":
            s.write_num(row, 4, value, 6 if bold else 5)

    def total_row(row, label, value):
        s.row_h(row, 18)
        s.write(row, 2, label, 7)
        s.write(row, 3, "",    7)
        s.write_num(row, 4, value, 6)

    # ── ASSETS ──
    section("ASSETS", r); r += 1
    section("  Current Assets", r); r += 1
    data_row(r, "  Cash & Cash Equivalents",        "~5-8% of revenue; liquidity buffer",          623);  r += 1
    data_row(r, "  Accounts Receivable",             "Franchise royalties & fees receivable",        312);  r += 1
    data_row(r, "  Inventory",                       "Food, packaging, supplies (~1-2 days)",         48);  r += 1
    data_row(r, "  Prepaid Expenses & Other",        "Insurance, rent prepayments",                   97);  r += 1
    total_row(r, "  Total Current Assets",          1080); r += 1
    r += 1
    section("  Non-Current Assets", r); r += 1
    data_row(r, "  Property, Plant & Equipment (Net)","~18,000 locations; owned + company-operated",2840); r += 1
    data_row(r, "  Right-of-Use Assets (Operating)", "IFRS 16 / ASC 842 lease capitalization",      3150); r += 1
    data_row(r, "  Intangible Assets",               "Franchise agreements, brand licenses",         1720); r += 1
    data_row(r, "  Goodwill",                        "Acquisition premiums (Popeyes, Firehouse)",    5580); r += 1
    data_row(r, "  Other Non-Current Assets",        "Equity investments, long-term deposits",        430); r += 1
    total_row(r, "  Total Non-Current Assets",      13720); r += 1
    r += 1
    total_row(r, "TOTAL ASSETS",                    14800); r += 2

    # ── LIABILITIES ──
    section("LIABILITIES", r); r += 1
    section("  Current Liabilities", r); r += 1
    data_row(r, "  Accounts Payable",               "Supplier terms ~30-45 days",                    287); r += 1
    data_row(r, "  Accrued Liabilities",             "Wages, marketing, utilities",                   412); r += 1
    data_row(r, "  Short-Term Lease Obligations",    "Current portion of lease liabilities",          298); r += 1
    data_row(r, "  Taxes Payable",                   "Corporate & deferred current portion",          143); r += 1
    data_row(r, "  Other Current Liabilities",       "Deferred franchise fees, gift cards",           195); r += 1
    total_row(r, "  Total Current Liabilities",     1335); r += 1
    r += 1
    section("  Non-Current Liabilities", r); r += 1
    data_row(r, "  Long-Term Debt",                  "Term loans, senior notes (3G legacy LBO)",     8420); r += 1
    data_row(r, "  Long-Term Lease Liabilities",     "15-20 year lease obligations",                 2890); r += 1
    data_row(r, "  Deferred Tax Liabilities",        "Timing differences on PP&E, goodwill",          680); r += 1
    data_row(r, "  Other Non-Current Liabilities",   "Pension obligations, deferred revenue",         225); r += 1
    total_row(r, "  Total Non-Current Liabilities", 12215); r += 1
    r += 1
    total_row(r, "TOTAL LIABILITIES",              13550); r += 2

    # ── EQUITY ──
    section("SHAREHOLDERS' EQUITY", r); r += 1
    data_row(r, "  Share Capital",                   "Common shares issued & outstanding",              3); r += 1
    data_row(r, "  Additional Paid-in Capital",      "IPO + secondary issuance premiums",            3840); r += 1
    data_row(r, "  Retained Earnings (Deficit)",     "Cumulative losses post-LBO / distributions", -2180); r += 1
    data_row(r, "  Accum. Other Comprehensive Income","FX translation, pension adjustments",         -413); r += 1
    total_row(r, "  Total Shareholders' Equity",    1250); r += 1
    r += 1
    total_row(r, "TOTAL LIABILITIES + EQUITY",     14800); r += 1

    # Balance check
    s.row_h(r, 20)
    s.write(r, 2, "✅  Balance Sheet Checks: Assets ($14,800M) = Liabilities ($13,550M) + Equity ($1,250M)", 13)
    s.merge(r, 2, r, 4)

    return s


# ─────────────────────────────────────────────────────────────────────────────
# SHEET 2 — FINANCIAL RATIOS & HEALTH
# ─────────────────────────────────────────────────────────────────────────────
def build_ratios():
    s = Sheet("Financial Ratios")
    s.col_w(1, 4); s.col_w(2, 28); s.col_w(3, 28); s.col_w(4, 14); s.col_w(5, 14); s.col_w(6, 14); s.col_w(7, 22)

    r = 1
    s.row_h(r, 40)
    s.write(r, 1, "BURGER KING — Financial Ratios & Health Analysis", 1)
    s.merge(r, 1, r, 7)
    r += 1
    s.write(r, 1, "FY2025E  |  All ratios based on estimated balance sheet figures", 15)
    s.merge(r, 1, r, 7); r += 2

    def hdr(row, labels):
        s.row_h(row, 22)
        for i, l in enumerate(labels):
            s.write(row, i+1, l, 3)

    def ratio_row(row, cat, name, formula, bk_val, benchmark, xf_val, comment):
        s.write(row, 1, cat,       4)
        s.write(row, 2, name,      4)
        s.write(row, 3, formula,   8)
        s.write(row, 4, bk_val,    12)
        s.write(row, 5, benchmark, 12)
        s.write(row, 6, "",        xf_val)
        s.write(row, 7, comment,   8)

    # ── Liquidity ──
    s.row_h(r, 20); s.write(r, 1, "A.  LIQUIDITY RATIOS", 2); s.merge(r, 1, r, 7); r += 1
    hdr(r, ["Category","Ratio","Formula","BK Value","Benchmark","Signal","Comment"]); r += 1
    ratio_row(r, "Liquidity", "Current Ratio",  "CA / CL",                    "0.81x","≥1.0x", 10, "Sub-par but typical for QSR franchisors"); r+=1
    ratio_row(r, "Liquidity", "Quick Ratio",    "(CA - Inventory) / CL",      "0.77x","≥0.8x", 10, "Thin; offset by strong OCF generation"); r+=1
    ratio_row(r, "Liquidity", "Cash Ratio",     "Cash / CL",                  "0.47x","≥0.3x",  9, "Adequate immediate liquidity cover"); r+=1
    r += 1

    # ── Leverage ──
    s.row_h(r, 20); s.write(r, 1, "B.  LEVERAGE & SOLVENCY RATIOS", 2); s.merge(r, 1, r, 7); r += 1
    hdr(r, ["Category","Ratio","Formula","BK Value","Benchmark","Signal","Comment"]); r += 1
    ratio_row(r, "Leverage",  "Debt-to-Equity",     "Total LT Debt / Equity",   "6.7x", "≤3.0x", 11, "Highly leveraged — 3G Capital legacy LBO"); r+=1
    ratio_row(r, "Leverage",  "Net Debt / EBITDA",  "Net Debt / EBITDA ~$2.1B", "~3.9x","≤4.0x", 10, "Approaching covenant thresholds"); r+=1
    ratio_row(r, "Leverage",  "Interest Coverage",  "EBIT / Interest Expense",  "~2.8x","≥3.0x", 10, "Borderline; vulnerable to rate rises"); r+=1
    ratio_row(r, "Leverage",  "Equity Multiplier",  "Total Assets / Equity",    "11.8x","~8x",   11, "Aggressive — minimal equity cushion"); r+=1
    r += 1

    # ── Asset Quality ──
    s.row_h(r, 20); s.write(r, 1, "C.  ASSET QUALITY", 2); s.merge(r, 1, r, 7); r += 1
    hdr(r, ["Category","Component","Description","USD (M)","% of Assets","Signal","Comment"]); r += 1
    ratio_row(r, "Asset Quality", "Goodwill + Intangibles", "Brand & franchise acquisition premiums", "$7,300","49.3%", 10, "Impairment risk if unit economics weaken"); r+=1
    ratio_row(r, "Asset Quality", "ROU Assets",             "IFRS 16 capitalized leases",             "$3,150","21.3%", 10, "Fixed obligation regardless of revenue"); r+=1
    ratio_row(r, "Asset Quality", "PP&E (Net)",             "Restaurant & kitchen assets",            "$2,840","19.2%",  9, "Moderate — mostly franchise model"); r+=1
    ratio_row(r, "Asset Quality", "Current Assets",         "Lean working capital",                   "$1,080"," 7.3%",  9, "Typical for asset-light QSR model"); r+=1
    r += 1

    # ── Scorecard ──
    s.row_h(r, 20); s.write(r, 1, "D.  OVERALL HEALTH SCORECARD", 2); s.merge(r, 1, r, 7); r += 1
    hdr(r, ["","Dimension","","Rating","","","Commentary"]); r += 1

    def score(row, dim, rating, xf_r, comment):
        s.write(row, 2, dim,     4)
        s.write(row, 4, rating,  xf_r)
        s.write(row, 7, comment, 8)

    score(r, "Liquidity",         "⚠️  FAIR",     10, "Adequate for franchise model cash dynamics"); r+=1
    score(r, "Leverage",          "🔴 HIGH",      11, "3G legacy capital structure; rate-sensitive"); r+=1
    score(r, "Asset Quality",     "⚠️  MODERATE", 10, "Goodwill-heavy; brand impairment risk"); r+=1
    score(r, "Solvency",          "⚠️  ADEQUATE", 10, "Cash flows cover obligations — for now"); r+=1
    score(r, "Earnings Quality",  "✅ GOOD",       9, "Royalty model = highly predictable revenue"); r+=1
    score(r, "Overall",           "⚠️  MODERATE", 10, "Viable but sensitive to rate & same-store sales"); r+=1

    return s


# ─────────────────────────────────────────────────────────────────────────────
# SHEET 3 — BK vs. McDONALD'S COMPARISON
# ─────────────────────────────────────────────────────────────────────────────
def build_comparison():
    s = Sheet("BK vs McDonalds")
    s.col_w(1, 4); s.col_w(2, 32); s.col_w(3, 18); s.col_w(4, 18); s.col_w(5, 36)

    r = 1
    s.row_h(r, 40)
    s.write(r, 1, "COMPARATIVE ANALYSIS — BURGER KING vs. McDONALD'S", 1)
    s.merge(r, 1, r, 5)
    r += 1
    s.write(r, 1, "Balance Sheet Structure Comparison  |  FY2025E  |  USD in Millions", 15)
    s.merge(r, 1, r, 5); r += 2

    s.row_h(r, 22)
    for c, h in [(2,"Metric"),(3,"Burger King (RBI)"),(4,"McDonald's (MCD)"),(5,"Commentary")]:
        s.write(r, c, h, 3)
    r += 1

    def comp_row(row, metric, bk, mcd, comment, xf_bk=0, xf_mcd=0):
        s.write(row, 2, metric,  4)
        s.write(row, 3, bk,      xf_bk  if xf_bk  else 12)
        s.write(row, 4, mcd,     xf_mcd if xf_mcd else 12)
        s.write(row, 5, comment, 8)

    rows_data = [
        ("Total Assets (USD)",         "~$14.8B",      "~$56.0B",     "MCD 4x larger; real estate portfolio drives asset base"),
        ("Total LT Debt",              "~$8.4B",       "~$37.0B",     "Both highly leveraged; MCD carries more in absolute terms"),
        ("Debt / EBITDA",              "~3.9x",        "~3.5x",       "MCD marginally more conservative on leverage ratio"),
        ("Goodwill + Intangibles",     "~49% of assets","~20% of assets","BK is acquisition-heavy; MCD owns real estate instead"),
        ("Book Equity",                "~$1.25B +ve",  "~-$7.0B -ve", "MCD has NEGATIVE equity from aggressive buyback programs"),
        ("Current Ratio",              "~0.81x",       "~1.0x",       "MCD slightly more liquid on current basis"),
        ("ROU / Lease Assets",         "~$3.2B",       "~$14.5B",     "MCD owns/leases far more properties as landlord"),
        ("Franchise Mix",              "~95% franchised","~95% franchised","Near-identical operating models at the restaurant level"),
        ("Revenue Model",              "Royalty + fees","Royalty + real estate rents","MCD's property leasing is the KEY strategic differentiator"),
        ("Goodwill Risk",              "HIGH (49% assets)","MODERATE (20%)","BK more exposed to intangible impairment charges"),
        ("Geographic Diversification", "100+ countries","120+ countries","Both highly global; similar FX translation exposure"),
        ("Capital Return Policy",      "Dividends + deleveraging","Buybacks dominant","MCD more aggressive on buybacks → negative book equity"),
    ]

    for item in rows_data:
        comp_row(r, *item); r += 1

    r += 1
    s.row_h(r, 20)
    s.write(r, 1, "KEY INSIGHT", 2); s.merge(r, 1, r, 5); r += 1
    insight = ("McDonald's operates as a real estate company disguised as a restaurant chain — owning land/buildings and leasing to franchisees. "
               "This creates a more defensible, diversified income stream vs. Burger King's pure royalty model. "
               "BK lacks the real estate moat, making it more exposed to franchisee concentration risk and commodity cycles.")
    s.write(r, 2, insight, 8); s.merge(r, 2, r, 5)
    s.row_h(r, 50)

    return s

# ─────────────────────────────────────────────────────────────────────────────
# SHEET 4 — RISK REGISTER
# ─────────────────────────────────────────────────────────────────────────────
def build_risks():
    s = Sheet("Risk Register")
    s.col_w(1, 4); s.col_w(2, 6); s.col_w(3, 30); s.col_w(4, 16); s.col_w(5, 16); s.col_w(6, 42)

    r = 1
    s.row_h(r, 40)
    s.write(r, 1, "BURGER KING — Key Risk Factors & Mitigation", 1)
    s.merge(r, 1, r, 6)
    r += 1
    s.write(r, 1, "Investment Risk Analysis  |  QSR Sector  |  FY2025E", 15)
    s.merge(r, 1, r, 6); r += 2

    s.row_h(r, 22)
    for c, h in [(2,"#"),(3,"Risk Factor"),(4,"Severity"),(5,"Likelihood"),(6,"Detail & Mitigation")]:
        s.write(r, c, h, 3)
    r += 1

    risks = [
        ("1","Debt Overhang & Refinancing Risk",       "🔴 CRITICAL","HIGH",
         "$8.4B LT debt; ~35% of EBIT consumed by interest. Refinancing risk if rates stay elevated. "
         "Mitigation: RBI has staggered maturity profile; strong OCF generation provides buffer."),
        ("2","Lease Liability Concentration",          "🟠 HIGH","HIGH",
         "$3.2B combined lease obligations are fixed outflows regardless of revenue. "
         "IFRS 16/ASC 842 recognition inflates reported liabilities vs pre-2019 comparisons."),
        ("3","Franchise Revenue Dependency",           "🟠 HIGH","MEDIUM",
         "95%+ of BK locations are franchised. Royalty income (~5% of system sales) is directly tied to "
         "franchisee health and same-store sales trends."),
        ("4","Goodwill & Intangible Impairment",       "🟡 MEDIUM","MEDIUM",
         "$5.58B goodwill from acquisition premiums. ASC 350 annual impairment test. "
         "Deteriorating unit economics could trigger non-cash write-downs impacting EPS."),
        ("5","Interest Rate Sensitivity",              "🟠 HIGH","HIGH",
         "Variable-rate debt exposure means EBIT coverage deteriorates as rates rise. "
         "Each +100bps move on floating debt adds ~$40-60M in annual interest expense."),
        ("6","FX & Geographic Risk",                  "🟡 MEDIUM","MEDIUM",
         "100+ country footprint creates translational FX risk. EM exposure "
         "(Latin America, Middle East) can distort reported royalty income materially."),
        ("7","Commodity & Supply Chain Inflation",     "🟡 MEDIUM","MEDIUM",
         "Beef, wheat, dairy, packaging costs directly impact franchisee margins and therefore "
         "system health. BK has limited direct exposure but franchisee stress flows through royalties."),
        ("8","Competitive Displacement",               "🟡 MEDIUM","LOW",
         "McDonald's, Wendy's, Chick-fil-A and fast-casual entrants (Shake Shack, etc.) "
         "compete for the same consumer. BK has underperformed MCD on same-store sales historically."),
    ]

    for risk in risks:
        num, name, severity, likelihood, detail = risk
        sev_xf = 11 if "CRITICAL" in severity else (10 if "HIGH" in severity else 9)
        s.write(r, 2, num,       4)
        s.write(r, 3, name,      4)
        s.write(r, 4, severity,  sev_xf)
        s.write(r, 5, likelihood,12)
        s.write(r, 6, detail,    8)
        s.row_h(r, 40)
        r += 1

    return s


# ─────────────────────────────────────────────────────────────────────────────
# SHEET 5 — ASSUMPTIONS & METHODOLOGY
# ─────────────────────────────────────────────────────────────────────────────
def build_assumptions():
    s = Sheet("Assumptions")
    s.col_w(1, 4); s.col_w(2, 32); s.col_w(3, 60)

    r = 1
    s.row_h(r, 40)
    s.write(r, 1, "ASSUMPTIONS & METHODOLOGY", 1)
    s.merge(r, 1, r, 3)
    r += 1
    s.write(r, 1, "Analytical Framework  |  QSR Industry Benchmarks  |  IFRS/US GAAP", 15)
    s.merge(r, 1, r, 3); r += 2

    s.row_h(r, 22)
    s.write(r, 2, "Assumption / Standard", 3); s.write(r, 3, "Detail", 3); r += 1

    items = [
        ("Revenue Base",             "BK system sales ~$12-13B globally; royalty rate ~5% = ~$600-650M fee revenue"),
        ("EBITDA Estimate",          "~$2.1B consolidated RBI EBITDA; BK segment ~$1.4-1.5B"),
        ("Cash Balance",             "~5-8% of revenue; management targets $500-700M minimum liquidity"),
        ("Accounts Receivable",      "Franchise royalties typically collected within 30-60 days; DSO ~17 days"),
        ("Inventory",                "QSR inventory turns >200x/year; minimal balance (~1-2 days of food cost)"),
        ("PP&E Basis",               "~18,000 BK locations globally; company-operated ~5%; rest franchised"),
        ("Right-of-Use Assets",      "IFRS 16 / ASC 842 adopted 2019; lease terms typically 15-20 years"),
        ("Goodwill",                 "Reflects TDL acquisition (2014), Popeyes (2017), Firehouse Subs (2021) premiums"),
        ("Long-Term Debt",           "Based on RBI public filings; ~$8.4B senior secured notes + term loans"),
        ("Lease Liabilities",        "Discounted at ~4-5% weighted average incremental borrowing rate"),
        ("Deferred Tax",             "Reflects accelerated depreciation on PP&E and franchise-related intangibles"),
        ("Shareholders' Equity",     "Negative retained earnings reflect LBO history, dividends, and impairment charges"),
        ("Accounting Standard",      "Hybrid IFRS/US GAAP presentation; RBI files 20-F under US GAAP with IFRS notes"),
        ("Comparable Company",       "McDonald's data sourced from MCD 10-K FY2024 public filings and analyst consensus"),
        ("Currency",                 "All figures in USD millions; non-USD revenues translated at period-average FX rates"),
        ("Disclaimer",               "Figures are analytical estimates for educational purposes. Not investment advice."),
    ]

    for assumption, detail in items:
        s.write(r, 2, assumption, 4)
        s.write(r, 3, detail,     8)
        s.row_h(r, 22)
        r += 1

    return s

# ─────────────────────────────────────────────────────────────────────────────
# ASSEMBLE THE XLSX
# ─────────────────────────────────────────────────────────────────────────────
sheets = [
    build_balance_sheet(),
    build_ratios(),
    build_comparison(),
    build_risks(),
    build_assumptions(),
]

CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml"  ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml"
    ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/sharedStrings.xml"
    ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
  <Override PartName="/xl/styles.xml"
    ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
""" + "".join(
    f'  <Override PartName="/xl/worksheets/sheet{i+1}.xml"\n'
    f'    ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>\n'
    for i in range(len(sheets))
) + "</Types>"

RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""

def wb_rels():
    lines = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
    ]
    for i in range(len(sheets)):
        lines.append(f'  <Relationship Id="rId{i+1}" '
                     f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
                     f'Target="worksheets/sheet{i+1}.xml"/>')
    lines.append(f'  <Relationship Id="rId{len(sheets)+1}" '
                 f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" '
                 f'Target="sharedStrings.xml"/>')
    lines.append(f'  <Relationship Id="rId{len(sheets)+2}" '
                 f'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" '
                 f'Target="styles.xml"/>')
    lines.append("</Relationships>")
    return "\n".join(lines)

def workbook_xml():
    sheets_el = "".join(
        f'<sheet name="{esc(s.name)}" sheetId="{i+1}" r:id="rId{i+1}"/>'
        for i, s in enumerate(sheets)
    )
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>{sheets_el}</sheets>
</workbook>"""

def sst_xml():
    items = "".join(f"<si><t>{esc(v)}</t></si>" for v in sst_list)
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            f'count="{len(sst_list)}" uniqueCount="{len(sst_list)}">{items}</sst>')

# ── trigger sheet XML generation (populates sst_list) ──
sheet_xmls = [s.to_xml() for s in sheets]

outfile = "/projects/sandbox/BurgerKing_BalanceSheet_Analysis.xlsx"
with zipfile.ZipFile(outfile, "w", zipfile.ZIP_DEFLATED) as zf:
    zf.writestr("[Content_Types].xml",             CONTENT_TYPES)
    zf.writestr("_rels/.rels",                     RELS)
    zf.writestr("xl/workbook.xml",                 workbook_xml())
    zf.writestr("xl/_rels/workbook.xml.rels",      wb_rels())
    zf.writestr("xl/styles.xml",                   STYLES_XML)
    zf.writestr("xl/sharedStrings.xml",            sst_xml())
    for i, xml in enumerate(sheet_xmls):
        zf.writestr(f"xl/worksheets/sheet{i+1}.xml", xml)

print(f"✅  Created: {outfile}  ({os.path.getsize(outfile):,} bytes)")
