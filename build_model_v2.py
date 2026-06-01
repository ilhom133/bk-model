#!/usr/bin/env python3
"""
Burger King (RBI segment) - 3-Statement Model v2 (+ Dashboard, Unit Economics, QSR KPIs)
Builds a real .xlsx (live formulas + cached values + styling) using only the stdlib,
and independently recomputes every figure to prove the balance sheet ties out.
"""
import zipfile, os

YEARS = [2024, 2025, 2026, 2027, 2028]
COL  = ['B', 'C', 'D', 'E', 'F']
PREV = ['A', 'B', 'C', 'D', 'E']

# ============================================================================
# 1) DRIVERS / ASSUMPTIONS
# ============================================================================
sss        = [0.030, 0.035, 0.030, 0.025, 0.025]   # same-store sales (= AUV growth)
unit_grow  = [None, 0.015, 0.015, 0.015, 0.015]    # net unit growth (2024 = base)
open_pct   = [0.030]*5                              # gross openings (% beg units)
close_pct  = [0.015]*5                              # closures (% beg units)
co_units   = [300, 270, 240, 220, 200]              # company-operated units (refranchising)
auv0       = 1.350                                  # 2024 system AUV ($M / unit)
royalty_r  = [0.045]*5
property_r = [0.015]*5
fee_r      = [0.0015]*5                             # franchise fees & other (% franchise sys sales)
cost_co_r  = [0.82]*5                               # cost of company restaurant sales
cost_fp_r  = [0.22]*5                               # franchise & property expense ratio
ga_r       = [0.13]*5                               # G&A (% of revenue)
da_r       = [0.06]*5                               # D&A (% of revenue)
tax_r      = [0.21]*5
int_r      = [0.05]*5
capex_ps   = [6000, 6200, 6400, 6600, 6800]         # capex per store ($)
lt_repay   = [200, 200, 200, 200, 200]              # LT debt repayment ($M)
payout     = [0.40]*5
shares     = [320.0]*5                              # shares (M, illustrative)

# ============================================================================
# 2) UNIT ECONOMICS
# ============================================================================
units = [19500.0]
for i in range(1, 5):
    units.append(units[-1]*(1+unit_grow[i]))
net_new  = [288.0] + [units[i]-units[i-1] for i in range(1, 5)]
openings = [576.0] + [units[i-1]*open_pct[i]  for i in range(1, 5)]
closures = [288.0] + [units[i-1]*close_pct[i] for i in range(1, 5)]
fr_units = [units[i]-co_units[i] for i in range(5)]
pct_fr   = [fr_units[i]/units[i] for i in range(5)]

auv = [auv0]
for i in range(1, 5):
    auv.append(auv[-1]*(1+sss[i]))
sys_sales = [auv[i]*units[i]    for i in range(5)]
co_sales  = [auv[i]*co_units[i] for i in range(5)]
fr_sales  = [auv[i]*fr_units[i] for i in range(5)]
sys_grow  = [None] + [sys_sales[i]/sys_sales[i-1]-1 for i in range(1, 5)]
rest_cost = [co_sales[i]*cost_co_r[i] for i in range(5)]
rest_eb   = [co_sales[i]-rest_cost[i] for i in range(5)]
rest_marg = [rest_eb[i]/co_sales[i] for i in range(5)]

# ============================================================================
# 3) INCOME STATEMENT
# ============================================================================
royalty  = [fr_sales[i]*royalty_r[i]  for i in range(5)]
prop_rev = [fr_sales[i]*property_r[i] for i in range(5)]
fees     = [fr_sales[i]*fee_r[i]      for i in range(5)]
revenue  = [co_sales[i]+royalty[i]+prop_rev[i]+fees[i] for i in range(5)]

cost_co  = [co_sales[i]*cost_co_r[i] for i in range(5)]
cost_fp  = [(royalty[i]+prop_rev[i])*cost_fp_r[i] for i in range(5)]
ga       = [revenue[i]*ga_r[i] for i in range(5)]
opex     = [cost_co[i]+cost_fp[i]+ga[i] for i in range(5)]
ebitda   = [revenue[i]-opex[i] for i in range(5)]
eb_marg  = [ebitda[i]/revenue[i] for i in range(5)]
da       = [revenue[i]*da_r[i] for i in range(5)]
ebit     = [ebitda[i]-da[i] for i in range(5)]

st_debt = [300.0]*5
lt_debt = [4700.0]
for i in range(1, 5):
    lt_debt.append(lt_debt[-1]-lt_repay[i])
tot_debt = [st_debt[i]+lt_debt[i] for i in range(5)]
interest = [int_r[i]*tot_debt[i] for i in range(5)]
pretax   = [ebit[i]-interest[i] for i in range(5)]
taxes    = [pretax[i]*tax_r[i] for i in range(5)]
ni       = [pretax[i]-taxes[i] for i in range(5)]
eps      = [ni[i]/shares[i] for i in range(5)]

# ============================================================================
# 4) BALANCE SHEET + CASH FLOW (integrated)
# ============================================================================
ar      = [0.07*revenue[i] for i in range(5)]
inv     = [0.01*revenue[i] for i in range(5)]
prepaid = [0.02*revenue[i] for i in range(5)]
ap      = [0.05*revenue[i] for i in range(5)]
accrued = [0.06*revenue[i] for i in range(5)]

rou, intang, gw       = [1400.0]*5, [1200.0]*5, [5500.0]*5
cur_lease, lease_l, dtl = [160.0]*5, [1300.0]*5, [350.0]*5
share_cap, apic, treas  = [100.0]*5, [2000.0]*5, [-600.0]*5

capex = [-(capex_ps[i]*units[i])/1_000_000 for i in range(5)]
ppe = [1500.0]
for i in range(1, 5):
    ppe.append(ppe[-1] + (capex_ps[i]*units[i])/1_000_000 - da[i])

dividends = [-payout[i]*ni[i] for i in range(5)]

CASH0 = 1000.0
tca0  = CASH0+ar[0]+inv[0]+prepaid[0]
tnca0 = ppe[0]+rou[0]+intang[0]+gw[0]
ta0   = tca0+tnca0
tcl0  = ap[0]+accrued[0]+st_debt[0]+cur_lease[0]
tncl0 = lt_debt[0]+lease_l[0]+dtl[0]
tl0   = tcl0+tncl0
re0   = ta0-tl0-(share_cap[0]+apic[0]+treas[0])
re = [re0]
for i in range(1, 5):
    re.append(re[-1]+ni[i]+dividends[i])

rev_23 = revenue[0]/1.05
wc23 = (0.07*rev_23, 0.01*rev_23, 0.02*rev_23, 0.05*rev_23, 0.06*rev_23)
chg_wc = [-((ar[0]-wc23[0])+(inv[0]-wc23[1])+(prepaid[0]-wc23[2]))
          + ((ap[0]-wc23[3])+(accrued[0]-wc23[4]))]
for i in range(1, 5):
    chg_wc.append(-((ar[i]-ar[i-1])+(inv[i]-inv[i-1])+(prepaid[i]-prepaid[i-1]))
                  + ((ap[i]-ap[i-1])+(accrued[i]-accrued[i-1])))

ocf = [ni[i]+da[i]+chg_wc[i] for i in range(5)]
icf = capex[:]
debt_flow = [-200.0] + [lt_debt[i]-lt_debt[i-1] for i in range(1, 5)]
fin_cf = [debt_flow[i]+dividends[i] for i in range(5)]
net_chg = [ocf[i]+icf[i]+fin_cf[i] for i in range(5)]
open_cash = [CASH0-net_chg[0]]
close_cash = [open_cash[0]+net_chg[0]]
for i in range(1, 5):
    open_cash.append(close_cash[-1]); close_cash.append(open_cash[i]+net_chg[i])
cash = close_cash[:]

tca  = [cash[i]+ar[i]+inv[i]+prepaid[i] for i in range(5)]
tnca = [ppe[i]+rou[i]+intang[i]+gw[i] for i in range(5)]
ta   = [tca[i]+tnca[i] for i in range(5)]
tcl  = [ap[i]+accrued[i]+st_debt[i]+cur_lease[i] for i in range(5)]
tncl = [lt_debt[i]+lease_l[i]+dtl[i] for i in range(5)]
tl   = [tcl[i]+tncl[i] for i in range(5)]
te   = [share_cap[i]+re[i]+apic[i]+treas[i] for i in range(5)]
tle  = [tl[i]+te[i] for i in range(5)]
check = [ta[i]-tle[i] for i in range(5)]
fcf = [ocf[i]+capex[i] for i in range(5)]

# ============================================================================
# VERIFY
# ============================================================================
print("="*88)
print("BURGER KING (RBI) - MODEL v2 VERIFICATION  (USD millions)")
print("="*88)
def show(lbl, arr, f="{:>12.2f}"):
    print(("{:<26}"+f*5).format(lbl, *arr))
print(("{:<26}"+"{:>12}"*5).format("Line", *[str(y) for y in YEARS]))
show("System-Wide Sales", sys_sales); show("Total Revenue", revenue)
show("Adj. EBITDA", ebitda); show("EBITDA Margin %", [m*100 for m in eb_marg])
show("Net Income", ni); show("Total Units", units, f="{:>12.0f}")
show("Free Cash Flow", fcf); show("Closing Cash", close_cash)
show("Total Debt", tot_debt); show("TOTAL ASSETS", ta); show("TOTAL L+E", tle)
print("-"*88)
show("BALANCE CHECK", check, f="{:>12.6f}")
ok = all(abs(c) < 1e-6 for c in check)
print("-"*88)
print("Debt/EBITDA   :", ["{:.2f}x".format(tot_debt[i]/ebitda[i]) for i in range(5)])
print("NetDebt/EBITDA:", ["{:.2f}x".format((tot_debt[i]-cash[i])/ebitda[i]) for i in range(5)])
print("ROE %         :", ["{:.1%}".format(ni[i]/te[i]) for i in range(5)])
print("Rest-Lvl Marg :", ["{:.1%}".format(rest_marg[i]) for i in range(5)])
print("EPS ($)       :", ["{:.2f}".format(eps[i]) for i in range(5)])
print("="*88)
print("BALANCE SHEET BALANCES IN ALL YEARS:", ok)
assert ok, "MODEL DOES NOT BALANCE"

def divisor(arr): return round(max(arr)/25.0, 2)
DIV = {"rev": divisor(revenue), "eb": divisor(ebitda), "sys": divisor(sys_sales),
       "fcf": divisor(fcf), "units": divisor(units)}

# ============================================================================
# 5) XLSX WRITER
# ============================================================================
def esc(s):
    return (str(s).replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;").replace('"', "&quot;"))
def numfmt(v):
    if v is None: return ""
    s = f"{v:.6f}".rstrip("0").rstrip(".")
    return "0" if s in ("", "-0") else s

NUMFMTS = {164:'#,##0.0',165:'0.0%',166:'#,##0',167:'0.0"x"',
           168:'"$"#,##0',169:'"$"#,##0.00',170:'#,##0.000'}
FONTS = [
    '<font><sz val="11"/><name val="Calibri"/></font>',
    '<font><b/><sz val="11"/><name val="Calibri"/></font>',
    '<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>',
    '<font><b/><sz val="16"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>',
    '<font><sz val="9"/><color rgb="FF808080"/><name val="Calibri"/></font>',
    '<font><b/><sz val="11"/><color rgb="FF2E75B6"/><name val="Consolas"/></font>',
]
FILLS = [
    '<fill><patternFill patternType="none"/></fill>',
    '<fill><patternFill patternType="gray125"/></fill>',
    '<fill><patternFill patternType="solid"><fgColor rgb="FF1F4E79"/></patternFill></fill>',
    '<fill><patternFill patternType="solid"><fgColor rgb="FF2E75B6"/></patternFill></fill>',
    '<fill><patternFill patternType="solid"><fgColor rgb="FFDDEBF7"/></patternFill></fill>',
]
XF_DEF = [
    ('lbl',0,0,0), ('lblb',0,1,0),
    ('num',164,0,0), ('numb',164,1,0),
    ('pct',165,0,0), ('pctb',165,1,0),
    ('int',166,0,0), ('intb',166,1,0),
    ('mult',167,0,0), ('multb',167,1,0),
    ('dol',168,0,0), ('dol2',169,0,0), ('auv',170,0,0),
    ('band',0,2,2), ('sect',0,2,3), ('title',0,3,2),
    ('numa',164,1,4), ('pcta',165,1,4), ('inta',166,1,4),
    ('dol2a',169,1,4), ('multa',167,1,4), ('lblba',0,1,4),
    ('note',0,4,0), ('bar',0,5,0),
]
STYLE = {name: i for i, (name, *_ ) in enumerate(XF_DEF)}

def styles_xml():
    nf = "".join(f'<numFmt numFmtId="{k}" formatCode="{esc(v)}"/>' for k, v in NUMFMTS.items())
    xfs = []
    for name, nfid, fid, flid in XF_DEF:
        a = ''
        if nfid >= 164: a += ' applyNumberFormat="1"'
        if fid != 0:    a += ' applyFont="1"'
        if flid > 1:    a += ' applyFill="1"'
        xfs.append(f'<xf numFmtId="{nfid}" fontId="{fid}" fillId="{flid}" borderId="0" xfId="0"{a}/>')
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            f'<numFmts count="{len(NUMFMTS)}">{nf}</numFmts>'
            f'<fonts count="{len(FONTS)}">{"".join(FONTS)}</fonts>'
            f'<fills count="{len(FILLS)}">{"".join(FILLS)}</fills>'
            '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
            '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
            f'<cellXfs count="{len(XF_DEF)}">{"".join(xfs)}</cellXfs>'
            '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
            '</styleSheet>')

class Cell:
    __slots__=("kind","val","cache","style","isstr")
    def __init__(s,kind,val=None,cache=None,style='lbl',isstr=False):
        s.kind,s.val,s.cache,s.style,s.isstr=kind,val,cache,style,isstr
def T(t,st='lbl'):  return Cell('s',t,style=st)
def N(v,st='num'):  return Cell('n',v,style=st)
def Fm(f,c,st='num',isstr=False): return Cell('f',f,c,style=st,isstr=isstr)
def E(st='lbl'):    return Cell('e',style=st)
def B():            return Cell('blank')

def cell_xml(ref, c):
    if c.kind=='blank': return ''
    s=STYLE[c.style]
    if c.kind=='e':   return f'<c r="{ref}" s="{s}"/>'
    if c.kind=='s':   return f'<c r="{ref}" s="{s}" t="inlineStr"><is><t xml:space="preserve">{esc(c.val)}</t></is></c>'
    if c.kind=='n':   return f'<c r="{ref}" s="{s}"><v>{numfmt(c.val)}</v></c>'
    if c.kind=='f':
        if c.isstr:   return f'<c r="{ref}" s="{s}" t="str"><f>{esc(c.val)}</f><v>{esc(c.cache)}</v></c>'
        return f'<c r="{ref}" s="{s}"><f>{esc(c.val)}</f><v>{numfmt(c.cache)}</v></c>'
    return ''

def sheet_xml(rows, colspec, freeze="B2"):
    out=['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
         '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
         'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">']
    if freeze:
        out.append('<sheetViews><sheetView showGridLines="0" workbookViewId="0">'
                   f'<pane xSplit="1" ySplit="1" topLeftCell="{freeze}" activePane="bottomRight" state="frozen"/>'
                   '</sheetView></sheetViews>')
    else:
        out.append('<sheetViews><sheetView showGridLines="0" workbookViewId="0"/></sheetViews>')
    out.append('<sheetFormatPr defaultRowHeight="15"/>')
    out.append('<cols>'+''.join(
        f'<col min="{a}" max="{b}" width="{w}" customWidth="1"/>' for a,b,w in colspec)+'</cols>')
    out.append('<sheetData>')
    for ridx,row in enumerate(rows,1):
        parts=[cell_xml(f"{chr(64+ci)}{ridx}",c) for ci,c in enumerate(row,1)]
        parts=[p for p in parts if p]
        out.append(f'<row r="{ridx}">'+''.join(parts)+'</row>' if parts else f'<row r="{ridx}"/>')
    out.append('</sheetData></worksheet>')
    return ''.join(out)

def hdr(label):
    return [T(label,'band')]+[N(y,'band') for y in YEARS]
def sect(text,n=6):
    return [T(text,'sect')]+[E('sect') for _ in range(n-1)]
def _vstyle(fmt,bold):
    bmap={'num':'numb','pct':'pctb','int':'intb','mult':'multb','dol':'dol','dol2':'dol2','auv':'auv'}
    return (bmap[fmt] if bold else fmt)
def drow(label, vals, fmt='num', bold=False):
    st=_vstyle(fmt,bold)
    cells=[T(label,'lblb' if bold else 'lbl')]
    for v in vals:
        if v is None: cells.append(B())
        elif v[0]=='f': cells.append(Fm(v[1],v[2],st))
        else: cells.append(N(v[1],st))
    return cells
def F(f,c): return ('f',f,c)
def Nv(v):  return ('n',v)

# ============================================================================
# 6) BUILD SHEETS
# ============================================================================
sheets={}

# ---------- Assumptions ----------
A=[hdr("ASSUMPTIONS  (USD millions unless noted)"), sect("UNIT & SALES DRIVERS")]
A.append(drow("Same-Store Sales (SSS) %", [Nv(v) for v in sss], 'pct'))
A.append(drow("Net Unit Growth %", [None]+[Nv(v) for v in unit_grow[1:]], 'pct'))
A.append(drow("Gross Openings (% beg. units)", [Nv(v) for v in open_pct], 'pct'))
A.append(drow("Closures (% beg. units)", [Nv(v) for v in close_pct], 'pct'))
A.append(drow("Company-Operated Units (#)", [Nv(v) for v in co_units], 'int'))
auv_cells=[Nv(auv0)]+[F(f"{PREV[i]}8*(1+{COL[i]}3)", auv[i]) for i in range(1,5)]
A.append(drow("AUV - System ($M / unit)", auv_cells, 'auv'))
A.append(sect("PRICING / FRANCHISE ECONOMICS"))
A.append(drow("Royalty Rate %", [Nv(v) for v in royalty_r], 'pct'))
A.append(drow("Property Revenue Rate %", [Nv(v) for v in property_r], 'pct'))
A.append(drow("Franchise Fees & Other Rate %", [Nv(v) for v in fee_r], 'pct'))
A.append(drow("Cost of Company Sales %", [Nv(v) for v in cost_co_r], 'pct'))
A.append(drow("Franchise & Property Cost %", [Nv(v) for v in cost_fp_r], 'pct'))
A.append(drow("G&A % of Revenue", [Nv(v) for v in ga_r], 'pct'))
A.append(sect("CAPITAL & FINANCING"))
A.append(drow("Depreciation & Amort. % Rev", [Nv(v) for v in da_r], 'pct'))
A.append(drow("Tax Rate %", [Nv(v) for v in tax_r], 'pct'))
A.append(drow("Interest Rate %", [Nv(v) for v in int_r], 'pct'))
A.append(drow("CapEx per Store ($)", [Nv(v) for v in capex_ps], 'dol'))
A.append(drow("LT Debt Repayment ($M)", [Nv(v) for v in lt_repay], 'num'))
A.append(drow("Dividend Payout %", [Nv(v) for v in payout], 'pct'))
A.append(drow("Shares Outstanding (M)", [Nv(v) for v in shares], 'num'))
sheets["Assumptions"]=A   # rows: SSS3 unitgrow4 open5 close6 coUnits7 AUV8 | roy10 prop11 fee12 costco13 costfp14 ga15 | da17 tax18 int19 capex20 repay21 payout22 shares23

# ---------- UnitEcon ----------
U=[hdr("UNIT ECONOMICS  (store-level)"), sect("UNIT COUNTS")]
tu=[Nv(19500.0)]+[F(f"{PREV[i]}3*(1+Assumptions!{COL[i]}4)", units[i]) for i in range(1,5)]
U.append(drow("Total Units (Ending)", tu, 'int', bold=True))
U.append(drow("Net New Units", [Nv(288.0)]+[F(f"{COL[i]}3-{PREV[i]}3", net_new[i]) for i in range(1,5)],'int'))
U.append(drow("Gross Openings", [Nv(576.0)]+[F(f"{PREV[i]}3*Assumptions!{COL[i]}5", openings[i]) for i in range(1,5)],'int'))
U.append(drow("Closures", [Nv(288.0)]+[F(f"{PREV[i]}3*Assumptions!{COL[i]}6", closures[i]) for i in range(1,5)],'int'))
U.append(drow("Company-Operated Units", [F(f"Assumptions!{COL[i]}7", co_units[i]) for i in range(5)],'int'))
U.append(drow("Franchised Units", [F(f"{COL[i]}3-{COL[i]}7", fr_units[i]) for i in range(5)],'int'))
U.append(drow("% Franchised", [F(f"{COL[i]}8/{COL[i]}3", pct_fr[i]) for i in range(5)],'pct'))
U.append(sect("UNIT VOLUMES & SYSTEM SALES"))
U.append(drow("AUV - System ($M / unit)", [F(f"Assumptions!{COL[i]}8", auv[i]) for i in range(5)],'auv'))
U.append(drow("System-Wide Sales ($M)", [F(f"{COL[i]}3*{COL[i]}11", sys_sales[i]) for i in range(5)],'num',bold=True))
U.append(drow("Company Restaurant Sales ($M)", [F(f"{COL[i]}7*{COL[i]}11", co_sales[i]) for i in range(5)],'num'))
U.append(drow("Franchised System Sales ($M)", [F(f"{COL[i]}8*{COL[i]}11", fr_sales[i]) for i in range(5)],'num'))
U.append(drow("System Sales Growth %", [None]+[F(f"{COL[i]}12/{PREV[i]}12-1", sys_grow[i]) for i in range(1,5)],'pct'))
U.append(sect("STORE-LEVEL (4-WALL) P&L"))
U.append(drow("Avg Sales / Company Unit ($M)", [F(f"{COL[i]}13/{COL[i]}7", co_sales[i]/co_units[i]) for i in range(5)],'auv'))
U.append(drow("Restaurant-Level Costs ($M)", [F(f"{COL[i]}13*Assumptions!{COL[i]}13", rest_cost[i]) for i in range(5)],'num'))
U.append(drow("Restaurant-Level EBITDA ($M)", [F(f"{COL[i]}13-{COL[i]}18", rest_eb[i]) for i in range(5)],'num'))
U.append(drow("Restaurant-Level Margin %", [F(f"{COL[i]}19/{COL[i]}13", rest_marg[i]) for i in range(5)],'pct',bold=True))
sheets["UnitEcon"]=U  # rows: units3 netnew4 open5 close6 coUnits7 frUnits8 %fr9 | AUV11 sys12 cosales13 frsales14 sysgrow15 | avg17 rcost18 reb19 rmarg20

# ---------- IS ----------
I=[hdr("INCOME STATEMENT  (USD millions)"), sect("REVENUE")]
I.append(drow("Company Restaurant Sales", [F(f"UnitEcon!{COL[i]}13", co_sales[i]) for i in range(5)],'num'))
I.append(drow("Franchise Royalties", [F(f"UnitEcon!{COL[i]}14*Assumptions!{COL[i]}10", royalty[i]) for i in range(5)],'num'))
I.append(drow("Property Revenue", [F(f"UnitEcon!{COL[i]}14*Assumptions!{COL[i]}11", prop_rev[i]) for i in range(5)],'num'))
I.append(drow("Franchise Fees & Other", [F(f"UnitEcon!{COL[i]}14*Assumptions!{COL[i]}12", fees[i]) for i in range(5)],'num'))
I.append(drow("Total Revenue", [F(f"SUM({COL[i]}3:{COL[i]}6)", revenue[i]) for i in range(5)],'num',bold=True))
I.append(sect("OPERATING COSTS"))
I.append(drow("Cost of Company Restaurant Sales", [F(f"{COL[i]}3*Assumptions!{COL[i]}13", cost_co[i]) for i in range(5)],'num'))
I.append(drow("Franchise & Property Expenses", [F(f"({COL[i]}4+{COL[i]}5)*Assumptions!{COL[i]}14", cost_fp[i]) for i in range(5)],'num'))
I.append(drow("General & Administrative (G&A)", [F(f"{COL[i]}7*Assumptions!{COL[i]}15", ga[i]) for i in range(5)],'num'))
I.append(drow("Total Operating Costs", [F(f"SUM({COL[i]}9:{COL[i]}11)", opex[i]) for i in range(5)],'num',bold=True))
I.append(drow("Adjusted EBITDA", [F(f"{COL[i]}7-{COL[i]}12", ebitda[i]) for i in range(5)],'num',bold=True))
I.append(drow("EBITDA Margin %", [F(f"{COL[i]}13/{COL[i]}7", eb_marg[i]) for i in range(5)],'pct'))
I.append(drow("Depreciation & Amortization", [F(f"{COL[i]}7*Assumptions!{COL[i]}17", da[i]) for i in range(5)],'num'))
I.append(drow("EBIT", [F(f"{COL[i]}13-{COL[i]}15", ebit[i]) for i in range(5)],'num',bold=True))
I.append(drow("Interest Expense", [F(f"Assumptions!{COL[i]}19*(BS!{COL[i]}17+BS!{COL[i]}20)", interest[i]) for i in range(5)],'num'))
I.append(drow("Pre-Tax Income", [F(f"{COL[i]}16-{COL[i]}17", pretax[i]) for i in range(5)],'num'))
I.append(drow("Taxes", [F(f"{COL[i]}18*Assumptions!{COL[i]}18", taxes[i]) for i in range(5)],'num'))
I.append(drow("Net Income", [F(f"{COL[i]}18-{COL[i]}19", ni[i]) for i in range(5)],'num',bold=True))
I.append(drow("EPS ($)", [F(f"{COL[i]}20/Assumptions!{COL[i]}23", eps[i]) for i in range(5)],'dol2'))
sheets["IS"]=I  # rows: cosales3 roy4 prop5 fee6 rev7 | costco9 costfp10 ga11 opex12 ebitda13 ebmarg14 da15 ebit16 int17 pretax18 tax19 ni20 eps21

# ---------- BS ----------
def constrow(label,base,vals,rn):
    return drow(label,[Nv(base)]+[F(f"{PREV[i]}{rn}", vals[i]) for i in range(1,5)],'num')
def wcrow(label,base,vals,rn):
    return drow(label,[Nv(base)]+[F(f"{PREV[i]}{rn}*IS!{COL[i]}7/IS!{PREV[i]}7", vals[i]) for i in range(1,5)],'num')
Bs=[hdr("BALANCE SHEET  (USD millions)"), sect("ASSETS")]
Bs.append(drow("Cash", [F(f"CF!{COL[i]}13", cash[i]) for i in range(5)],'num'))
Bs.append(wcrow("Accounts Receivable",ar[0],ar,4))
Bs.append(wcrow("Inventory",inv[0],inv,5))
Bs.append(wcrow("Prepaid Expenses",prepaid[0],prepaid,6))
Bs.append(drow("Total Current Assets",[F(f"SUM({COL[i]}3:{COL[i]}6)", tca[i]) for i in range(5)],'num',bold=True))
ppe_cells=[Nv(1500.0)]+[F(f"{PREV[i]}8+Assumptions!{COL[i]}20*UnitEcon!{COL[i]}3/1000000-IS!{COL[i]}15", ppe[i]) for i in range(1,5)]
Bs.append(drow("PP&E",ppe_cells,'num'))
Bs.append(constrow("Right-of-Use Assets",1400.0,rou,9))
Bs.append(constrow("Intangibles",1200.0,intang,10))
Bs.append(constrow("Goodwill",5500.0,gw,11))
Bs.append(drow("Total Non-Current Assets",[F(f"SUM({COL[i]}8:{COL[i]}11)", tnca[i]) for i in range(5)],'num',bold=True))
Bs.append(drow("TOTAL ASSETS",[F(f"{COL[i]}7+{COL[i]}12", ta[i]) for i in range(5)],'num',bold=True))
Bs.append(sect("LIABILITIES"))
Bs.append(wcrow("Accounts Payable",ap[0],ap,15))
Bs.append(wcrow("Accrued Expenses",accrued[0],accrued,16))
Bs.append(constrow("Short-term Debt",300.0,st_debt,17))
Bs.append(constrow("Current Lease Liability",160.0,cur_lease,18))
Bs.append(drow("Total Current Liabilities",[F(f"SUM({COL[i]}15:{COL[i]}18)", tcl[i]) for i in range(5)],'num',bold=True))
ltd_cells=[Nv(4700.0)]+[F(f"{PREV[i]}20-Assumptions!{COL[i]}21", lt_debt[i]) for i in range(1,5)]
Bs.append(drow("Long-term Debt",ltd_cells,'num'))
Bs.append(constrow("Lease Liabilities",1300.0,lease_l,21))
Bs.append(constrow("Deferred Tax",350.0,dtl,22))
Bs.append(drow("Total Non-Current Liabilities",[F(f"SUM({COL[i]}20:{COL[i]}22)", tncl[i]) for i in range(5)],'num',bold=True))
Bs.append(drow("TOTAL LIABILITIES",[F(f"{COL[i]}19+{COL[i]}23", tl[i]) for i in range(5)],'num',bold=True))
Bs.append(sect("EQUITY"))
Bs.append(constrow("Share Capital",100.0,share_cap,26))
re_cells=[Nv(re0)]+[F(f"{PREV[i]}27+IS!{COL[i]}20+CF!{COL[i]}9", re[i]) for i in range(1,5)]
Bs.append(drow("Retained Earnings",re_cells,'num'))
Bs.append(constrow("Additional Paid-In Capital",2000.0,apic,28))
Bs.append(constrow("Treasury Stock",-600.0,treas,29))
Bs.append(drow("TOTAL EQUITY",[F(f"{COL[i]}26+{COL[i]}27+{COL[i]}28+{COL[i]}29", te[i]) for i in range(5)],'num',bold=True))
Bs.append(drow("TOTAL LIABILITIES + EQUITY",[F(f"{COL[i]}24+{COL[i]}30", tle[i]) for i in range(5)],'num',bold=True))
Bs.append(drow("Balance Check (must = 0)",[F(f"{COL[i]}13-{COL[i]}31", check[i]) for i in range(5)],'num',bold=True))
sheets["BS"]=Bs  # rows: cash3 ar4 inv5 prep6 tca7 ppe8 rou9 intang10 gw11 tnca12 TA13 | ap15 accr16 std17 curlease18 tcl19 ltd20 lease21 dtl22 tncl23 TL24 | sc26 re27 apic28 treas29 TE30 TLE31 chk32

# ---------- CF (no section rows -> references line up) ----------
Cf=[hdr("CASH FLOW STATEMENT  (USD millions)")]
Cf.append(drow("Net Income",[F(f"IS!{COL[i]}20", ni[i]) for i in range(5)],'num'))                       # 2
Cf.append(drow("Depreciation & Amortization",[F(f"IS!{COL[i]}15", da[i]) for i in range(5)],'num'))      # 3
wc=[Nv(chg_wc[0])]
for i in range(1,5):
    c,p=COL[i],PREV[i]
    wc.append(F(f"-((BS!{c}4-BS!{p}4)+(BS!{c}5-BS!{p}5)+(BS!{c}6-BS!{p}6))+((BS!{c}15-BS!{p}15)+(BS!{c}16-BS!{p}16))",chg_wc[i]))
Cf.append(drow("Change in Working Capital",wc,'num'))                                                    # 4
Cf.append(drow("Operating Cash Flow",[F(f"{COL[i]}2+{COL[i]}3+{COL[i]}4", ocf[i]) for i in range(5)],'num',bold=True))  # 5
Cf.append(drow("CapEx",[F(f"-Assumptions!{COL[i]}20*UnitEcon!{COL[i]}3/1000000", capex[i]) for i in range(5)],'num'))   # 6
Cf.append(drow("Investing Cash Flow",[F(f"{COL[i]}6", icf[i]) for i in range(5)],'num',bold=True))       # 7
db=[Nv(-200.0)]+[F(f"BS!{COL[i]}20-BS!{PREV[i]}20", debt_flow[i]) for i in range(1,5)]
Cf.append(drow("Debt Issued/(Repaid)",db,'num'))                                                          # 8
Cf.append(drow("Dividends",[F(f"-Assumptions!{COL[i]}22*IS!{COL[i]}20", dividends[i]) for i in range(5)],'num'))  # 9
Cf.append(drow("Financing Cash Flow",[F(f"{COL[i]}8+{COL[i]}9", fin_cf[i]) for i in range(5)],'num',bold=True))   # 10
Cf.append(drow("Net Change in Cash",[F(f"{COL[i]}5+{COL[i]}7+{COL[i]}10", net_chg[i]) for i in range(5)],'num',bold=True))  # 11
oc=[Nv(open_cash[0])]+[F(f"{PREV[i]}13", open_cash[i]) for i in range(1,5)]
Cf.append(drow("Opening Cash",oc,'num'))                                                                  # 12
Cf.append(drow("Closing Cash",[F(f"{COL[i]}12+{COL[i]}11", close_cash[i]) for i in range(5)],'num',bold=True))  # 13
Cf.append(drow("Interest Paid (memo; in Net Income)",[F(f"-IS!{COL[i]}17", -interest[i]) for i in range(5)],'num'))  # 14
sheets["CF"]=Cf

# ---------- KPIs ----------
K=[hdr("KEY QSR & FINANCIAL KPIs"), sect("GROWTH & UNITS")]
K.append(drow("System-Wide Sales ($M)",[F(f"UnitEcon!{COL[i]}12", sys_sales[i]) for i in range(5)],'num'))
K.append(drow("System Sales Growth %",[None]+[F(f"UnitEcon!{COL[i]}15", sys_grow[i]) for i in range(1,5)],'pct'))
K.append(drow("Total Units",[F(f"UnitEcon!{COL[i]}3", units[i]) for i in range(5)],'int'))
K.append(drow("Net New Units",[F(f"UnitEcon!{COL[i]}4", net_new[i]) for i in range(5)],'int'))
K.append(drow("Unit Growth %",[None]+[F(f"UnitEcon!{COL[i]}3/UnitEcon!{PREV[i]}3-1", units[i]/units[i-1]-1) for i in range(1,5)],'pct'))
K.append(drow("% Franchised",[F(f"UnitEcon!{COL[i]}9", pct_fr[i]) for i in range(5)],'pct'))
K.append(drow("AUV - System ($M)",[F(f"UnitEcon!{COL[i]}11", auv[i]) for i in range(5)],'auv'))
K.append(drow("Same-Store Sales %",[F(f"Assumptions!{COL[i]}3", sss[i]) for i in range(5)],'pct'))
K.append(sect("PROFITABILITY"))
K.append(drow("Total Revenue ($M)",[F(f"IS!{COL[i]}7", revenue[i]) for i in range(5)],'num'))
K.append(drow("Revenue Growth %",[None]+[F(f"IS!{COL[i]}7/IS!{PREV[i]}7-1", revenue[i]/revenue[i-1]-1) for i in range(1,5)],'pct'))
K.append(drow("Adjusted EBITDA ($M)",[F(f"IS!{COL[i]}13", ebitda[i]) for i in range(5)],'num'))
K.append(drow("EBITDA Margin %",[F(f"IS!{COL[i]}14", eb_marg[i]) for i in range(5)],'pct'))
K.append(drow("Restaurant-Level Margin %",[F(f"UnitEcon!{COL[i]}20", rest_marg[i]) for i in range(5)],'pct'))
K.append(drow("Net Income ($M)",[F(f"IS!{COL[i]}20", ni[i]) for i in range(5)],'num'))
K.append(drow("Net Margin %",[F(f"IS!{COL[i]}20/IS!{COL[i]}7", ni[i]/revenue[i]) for i in range(5)],'pct'))
K.append(drow("EPS ($)",[F(f"IS!{COL[i]}21", eps[i]) for i in range(5)],'dol2'))
K.append(drow("Royalty Revenue ($M)",[F(f"IS!{COL[i]}4", royalty[i]) for i in range(5)],'num'))
K.append(drow("Effective Royalty Rate %",[F(f"IS!{COL[i]}4/UnitEcon!{COL[i]}12", royalty[i]/sys_sales[i]) for i in range(5)],'pct'))
K.append(sect("RETURNS & EFFICIENCY"))
K.append(drow("ROE %",[F(f"IS!{COL[i]}20/BS!{COL[i]}30", ni[i]/te[i]) for i in range(5)],'pct'))
K.append(drow("ROA %",[F(f"IS!{COL[i]}20/BS!{COL[i]}13", ni[i]/ta[i]) for i in range(5)],'pct'))
K.append(drow("ROIC %",[F(f"IS!{COL[i]}16*(1-Assumptions!{COL[i]}18)/(BS!{COL[i]}17+BS!{COL[i]}20+BS!{COL[i]}30-BS!{COL[i]}3)",
              ebit[i]*(1-tax_r[i])/(tot_debt[i]+te[i]-cash[i])) for i in range(5)],'pct'))
K.append(drow("G&A % of System Sales",[F(f"IS!{COL[i]}11/UnitEcon!{COL[i]}12", ga[i]/sys_sales[i]) for i in range(5)],'pct'))
K.append(drow("CapEx ($M)",[F(f"-CF!{COL[i]}6", -capex[i]) for i in range(5)],'num'))
K.append(drow("CapEx % of Revenue",[F(f"-CF!{COL[i]}6/IS!{COL[i]}7", -capex[i]/revenue[i]) for i in range(5)],'pct'))
K.append(sect("CASH & LEVERAGE"))
K.append(drow("Operating Cash Flow ($M)",[F(f"CF!{COL[i]}5", ocf[i]) for i in range(5)],'num'))
K.append(drow("Free Cash Flow ($M)",[F(f"CF!{COL[i]}5+CF!{COL[i]}6", fcf[i]) for i in range(5)],'num',bold=True))
K.append(drow("FCF Conversion % (FCF/EBITDA)",[F(f"(CF!{COL[i]}5+CF!{COL[i]}6)/IS!{COL[i]}13", fcf[i]/ebitda[i]) for i in range(5)],'pct'))
K.append(drow("Total Debt ($M)",[F(f"BS!{COL[i]}17+BS!{COL[i]}20", tot_debt[i]) for i in range(5)],'num'))
K.append(drow("Net Debt ($M)",[F(f"BS!{COL[i]}17+BS!{COL[i]}20-BS!{COL[i]}3", tot_debt[i]-cash[i]) for i in range(5)],'num'))
K.append(drow("Debt / EBITDA (x)",[F(f"(BS!{COL[i]}17+BS!{COL[i]}20)/IS!{COL[i]}13", tot_debt[i]/ebitda[i]) for i in range(5)],'mult'))
K.append(drow("Net Debt / EBITDA (x)",[F(f"(BS!{COL[i]}17+BS!{COL[i]}20-BS!{COL[i]}3)/IS!{COL[i]}13", (tot_debt[i]-cash[i])/ebitda[i]) for i in range(5)],'mult'))
K.append(drow("Interest Coverage (x)",[F(f"IS!{COL[i]}16/IS!{COL[i]}17", ebit[i]/interest[i]) for i in range(5)],'mult'))
K.append(drow("Dividends ($M)",[F(f"-CF!{COL[i]}9", -dividends[i]) for i in range(5)],'num'))
K.append(drow("Dividend Payout %",[F(f"-CF!{COL[i]}9/IS!{COL[i]}20", -dividends[i]/ni[i]) for i in range(5)],'pct'))
K.append(drow("Dividend per Share ($)",[F(f"-CF!{COL[i]}9/Assumptions!{COL[i]}23", -dividends[i]/shares[i]) for i in range(5)],'dol2'))
sheets["KPIs"]=K

# ---------- Dashboard ----------
def cagr(b,e): return (e/b)**0.25-1
D=[]
D.append([T("BURGER KING (RBI SEGMENT) - 3-STATEMENT MODEL & QSR DASHBOARD",'title')]+[E('title') for _ in range(7)])
D.append([T("USD millions unless noted  |  2024A - 2028E  |  Fully integrated & balanced (BS check = 0)",'note')]+[B()]*7)
D.append([B()]*8)
D.append([T("HEADLINE KPIs",'sect')]+[E('sect') for _ in range(7)])
D.append([T("Metric",'band'),T("2024A",'band'),T("2028E",'band'),T("'24-'28",'band')]+[E('band') for _ in range(4)])
HEAD=[
 ("System-Wide Sales ($M)","UnitEcon!B12","UnitEcon!F12",sys_sales[0],sys_sales[4],'numa',False),
 ("Total Revenue ($M)","IS!B7","IS!F7",revenue[0],revenue[4],'numa',False),
 ("Adjusted EBITDA ($M)","IS!B13","IS!F13",ebitda[0],ebitda[4],'numa',False),
 ("Net Income ($M)","IS!B20","IS!F20",ni[0],ni[4],'numa',False),
 ("Free Cash Flow ($M)","CF!B5+CF!B6","CF!F5+CF!F6",fcf[0],fcf[4],'numa',False),
 ("Total Units","UnitEcon!B3","UnitEcon!F3",units[0],units[4],'inta',False),
 ("EBITDA Margin %","IS!B14","IS!F14",eb_marg[0],eb_marg[4],'pcta',True),
 ("Net Debt / EBITDA (x)","(BS!B17+BS!B20-BS!B3)/IS!B13","(BS!F17+BS!F20-BS!F3)/IS!F13",
      (tot_debt[0]-cash[0])/ebitda[0],(tot_debt[4]-cash[4])/ebitda[4],'multa',True),
 ("ROE %","IS!B20/BS!B30","IS!F20/BS!F30",ni[0]/te[0],ni[4]/te[4],'pcta',True),
]
r=6
for (lbl,bref,fref,b,e,fmt,delta) in HEAD:
    if delta:
        dfmt = 'multa' if fmt=='multa' else 'pcta'
        dcell=Fm(f"C{r}-B{r}", e-b, dfmt)
    else:
        dcell=Fm(f"(C{r}/B{r})^(1/4)-1", cagr(b,e), 'pcta')
    D.append([T(lbl,'lblba'),Fm(bref,b,fmt),Fm(fref,e,fmt),dcell]+[B()]*4)
    r+=1
D.append([B()]*8)
D.append([T("VISUAL TRENDS  (each \u2588 scaled to series max)",'sect')]+[E('sect') for _ in range(7)])
def bars_block(title, ref_tmpl, arr, div):
    D.append([T(title,'lblb')]+[E('lbl')]*2)
    for i in range(5):
        nb=int(round(arr[i]/div))
        D.append([N(YEARS[i],'lbl'),
                  Fm(ref_tmpl.format(c=COL[i]), arr[i], 'num'),
                  Fm(f'REPT("\u2588",ROUND(({ref_tmpl.format(c=COL[i])})/{div},0))', "\u2588"*nb, 'bar', isstr=True)])
bars_block("Total Revenue ($M)","IS!{c}7",revenue,DIV["rev"])
bars_block("Adjusted EBITDA ($M)","IS!{c}13",ebitda,DIV["eb"])
bars_block("System-Wide Sales ($M)","UnitEcon!{c}12",sys_sales,DIV["sys"])
bars_block("Free Cash Flow ($M)","CF!{c}5+CF!{c}6",fcf,DIV["fcf"])
bars_block("Total Units","UnitEcon!{c}3",units,DIV["units"])
D.append([B()]*8)
D.append([T("Tabs: Assumptions | UnitEcon | IS | BS | CF | KPIs.  2024 = base/actual; 2025-2028 = live formulas.",'note')]+[B()]*7)
sheets["Dashboard"]=D

# ============================================================================
# 7) WRITE WORKBOOK
# ============================================================================
ORDER=["Dashboard","Assumptions","UnitEcon","IS","BS","CF","KPIs"]
COLS_WIDE=[(1,1,34),(2,6,15)]
COLS_DASH=[(1,1,30),(2,2,14),(3,3,46),(4,8,12)]
COLSPEC={"Dashboard":COLS_DASH,"Assumptions":COLS_WIDE,"UnitEcon":COLS_WIDE,
         "IS":COLS_WIDE,"BS":COLS_WIDE,"CF":COLS_WIDE,"KPIs":COLS_WIDE}
FREEZE={"Dashboard":None,"Assumptions":"B2","UnitEcon":"B2","IS":"B2","BS":"B2","CF":"B2","KPIs":"B2"}

def write_xlsx(path):
    ct=['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">',
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
        '<Default Extension="xml" ContentType="application/xml"/>',
        '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>',
        '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>']
    for n in range(1,len(ORDER)+1):
        ct.append(f'<Override PartName="/xl/worksheets/sheet{n}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>')
    ct.append('</Types>')
    root_rels=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
               '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
               '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
               '</Relationships>')
    sx="".join(f'<sheet name="{esc(ORDER[i])}" sheetId="{i+1}" r:id="rId{i+1}"/>' for i in range(len(ORDER)))
    workbook=('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
              'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
              f'<sheets>{sx}</sheets><calcPr calcId="0" fullCalcOnLoad="1"/></workbook>')
    wb_rels=['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
             '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">']
    for n in range(1,len(ORDER)+1):
        wb_rels.append(f'<Relationship Id="rId{n}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{n}.xml"/>')
    wb_rels.append(f'<Relationship Id="rId{len(ORDER)+1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>')
    wb_rels.append('</Relationships>')
    with zipfile.ZipFile(path,"w",zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml","".join(ct))
        z.writestr("_rels/.rels",root_rels)
        z.writestr("xl/workbook.xml",workbook)
        z.writestr("xl/_rels/workbook.xml.rels","".join(wb_rels))
        z.writestr("xl/styles.xml",styles_xml())
        for i,name in enumerate(ORDER,1):
            z.writestr(f"xl/worksheets/sheet{i}.xml",sheet_xml(sheets[name],COLSPEC[name],FREEZE[name]))

OUT="/projects/sandbox/Burger_King_RBI_3_Statement_Model_v2.xlsx"
write_xlsx(OUT)
print("\nWrote:",OUT,"(",os.path.getsize(OUT),"bytes )")
