"""
build.py
AllView Real Estate — Dashboard Builder
Reads Excel files from ./data/, runs all calculations, outputs ./output/index.html

Run manually:  python scripts/build.py
In CI:         called automatically by GitHub Actions after download_files.py
"""

import os, re, sys, json, datetime
import openpyxl
from collections import defaultdict

DATA_DIR   = './data'
OUTPUT_DIR = './output'
os.makedirs(OUTPUT_DIR, exist_ok=True)

print("=" * 60)
print("AllView Dashboard Builder")
print("=" * 60)

# ── FILE PATHS ───────────────────────────────────────────────────────────────
FILES = {
    'portfolio':     os.path.join(DATA_DIR, 'New_Portafolios_Structure_-_Effective_Feb_17.xlsx'),
    'vacancy':       os.path.join(DATA_DIR, 'Rental_Vacancy.xlsx'),
    'applications':  os.path.join(DATA_DIR, 'rental_applications-Performance.xlsx'),
    'move_inout':    os.path.join(DATA_DIR, 'tenant_tickler_-_Move-in-out.xlsx'),
    'unpaid':        os.path.join(DATA_DIR, 'tenant_unpaid_charges_summary.xlsx'),
    'unit_turn':     os.path.join(DATA_DIR, 'Unit_Turn_Details.xlsx'),
    'work_orders':   os.path.join(DATA_DIR, 'WorkOrders_Report.xlsx'),
    'reviews_raw':   os.path.join(DATA_DIR, 'Yelp-Google_Reviews.xlsx'),
    'five_star':     os.path.join(DATA_DIR, '5-star_Tracker.xlsx'),
    'bdm':           os.path.join(DATA_DIR, 'BDM_Data.xlsx'),
    'email':         os.path.join(DATA_DIR, 'Email_Analytics.xlsx'),
    'guest_card':    os.path.join(DATA_DIR, 'guest_card_inquiries-Performance.xlsx'),
    'lr':            os.path.join(DATA_DIR, 'LR_Performance.xlsx'),
    'new_nps':       os.path.join(DATA_DIR, 'New_NPS_Data.xlsx'),
    'nps':           os.path.join(DATA_DIR, 'NPS_Data.xlsx'),
    'concessions':   os.path.join(DATA_DIR, 'Owner_Concessions.xlsx'),
    'phone':         os.path.join(DATA_DIR, 'Phone_Reporting.xlsx'),
    'rent_roll':     os.path.join(DATA_DIR, 'Rent_Roll.xlsx'),
}

def check_files():
    missing = [k for k, v in FILES.items() if not os.path.exists(v)]
    if missing:
        print(f"WARNING: Missing files: {missing}")
    else:
        print(f"[1/8] All {len(FILES)} source files found")

check_files()

# ── HELPERS ──────────────────────────────────────────────────────────────────
def read_sheet(path, sheet=0):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet] if isinstance(sheet, str) else wb.worksheets[sheet]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    return rows

def eom_we_map(rows, year_col, month_col, we_col):
    """Return {(year,month): last_WE_date} from a set of rows."""
    by_month = defaultdict(list)
    for r in rows:
        if not r or not r[we_col] or not hasattr(r[we_col],'strftime'): continue
        key = (str(r[year_col]).strip() if r[year_col] else '',
               str(r[month_col]).strip() if r[month_col] else '')
        by_month[key].append(r[we_col])
    return {key: max(dates) for key, dates in by_month.items()}

def agg_keys(year, month, ter, prop):
    y,m,t,p = year,month,ter,prop
    return [
        f"{y}|{m}|{t}|{p}", f"{y}|{m}|{t}|", f"{y}|{m}||{p}", f"{y}|{m}||",
        f"|{m}|{t}|{p}",    f"|{m}|{t}|",    f"|{m}||{p}",    f"|{m}||",
        f"{y}||{t}|{p}",    f"{y}||{t}|",    f"{y}|||{p}",    f"{y}|||",
        f"||{t}|{p}",       f"||{t}|",       f"|||{p}",       "|||",
    ]

MONTHS = ['January','February','March','April','May','June',
          'July','August','September','October','November','December']
SKIP_TERS = {'TBD', 'Mid OC', ''}

# ── PORTFOLIO MASTER RESOLVE ──────────────────────────────────────────────────
print("[2/8] Building portfolio master mapping...")
port_rows  = read_sheet(FILES['portfolio'])
master_full  = {}
master_short = defaultdict(list)
for r in port_rows[1:]:
    if not r or not r[0]: continue
    addr  = str(r[0]).strip()
    attrs = {
        'territory': str(r[6]).strip() if r[6] else '',
        'proptype':  str(r[5]).strip() if r[5] else '',
        'pm':        str(r[8]).strip() if r[8] else '',
        'la':        str(r[7]).strip() if r[7] else '',
    }
    master_full[addr] = attrs
    short = addr.split(',')[0].strip().lower()
    if short not in master_short: master_short[short] = [addr]
    if r[1]: master_short[str(r[1]).strip().lower()].append(addr)

def resolve(name):
    if not name: return None
    name = str(name).strip()
    if name in master_full: return master_full[name]
    short = name.split(',')[0].strip().lower()
    if short in master_short: return master_full[master_short[short][0]]
    base = re.split(r'\s*[\|#]\s*', name)[0].strip()
    base = re.sub(r'\s+(unit|apt|suite|ste)\s+\S+$','',base,flags=re.I).strip()
    short2 = base.split(',')[0].strip().lower()
    if short2 in master_short: return master_full[master_short[short2][0]]
    if ' - ' in name:
        full_part = name.split(' - ',1)[1].strip()
        if full_part in master_full: return master_full[full_part]
        short3 = full_part.split(',')[0].strip().lower()
        if short3 in master_short: return master_full[master_short[short3][0]]
    return None

# ── SECTION 3: BUILD DATA TABLES ─────────────────────────────────────────────
print("[3/8] Computing DATA tables (move-in/out, work orders, LR, NPS, unit turn)...")

DATA = {}

# ── MOVE IN / OUT ─────────────────────────────────────────────────────────────
mio_rows = read_sheet(FILES['move_inout'])
mi_tbl = defaultdict(int); mo_tbl = defaultdict(int)
mi_we  = defaultdict(int); mo_we  = defaultdict(int)
for r in mio_rows[1:]:
    if not r or not r[0]: continue
    attrs = resolve(r[0]) or {}
    ter   = attrs.get('territory','')
    prop  = attrs.get('proptype','')
    event = str(r[3]).strip() if r[3] else ''
    month = str(r[23]).strip() if r[23] else ''
    year  = str(r[22]).strip() if r[22] else ''
    we    = r[24].strftime('%Y-%m-%d') if r[24] and hasattr(r[24],'strftime') else ''
    tbl = mi_tbl if event=='Move-in' else mo_tbl if event=='Move-out' else None
    if tbl is None: continue
    for k in agg_keys(year,month,ter,prop): tbl[k] += 1
    # WE level
    we_tbl = mi_we if event=='Move-in' else mo_we
    if we:
        we_tbl[f"{we}|{ter}|{prop}"] += 1
        we_tbl[f"{we}|{ter}|"]       += 1
        we_tbl[f"{we}||"]            += 1

DATA['mi'] = dict(mi_tbl); DATA['mo'] = dict(mo_tbl)
DATA['mi_we'] = dict(mi_we); DATA['mo_we'] = dict(mo_we)

# ── WORK ORDERS ───────────────────────────────────────────────────────────────
wo_rows = read_sheet(FILES['work_orders'])
wo_tbl = defaultdict(int); wo_we_tbl = defaultdict(int)
for r in wo_rows[1:]:
    if not r or not r[0]: continue
    attrs = resolve(r[0]) or {}
    ter   = attrs.get('territory',''); prop = attrs.get('proptype','')
    month = str(r[43]).strip() if r[43] else ''
    year  = str(r[42]).strip() if r[42] else ''
    we    = r[44].strftime('%Y-%m-%d') if len(r)>44 and r[44] and hasattr(r[44],'strftime') else ''
    for k in agg_keys(year,month,ter,prop): wo_tbl[k] += 1
    if we:
        wo_we_tbl[f"{we}|{ter}|{prop}"] += 1
        wo_we_tbl[f"{we}|{ter}|"]       += 1
        wo_we_tbl[f"{we}||"]            += 1
DATA['wo'] = dict(wo_tbl); DATA['wo_we'] = dict(wo_we_tbl)

# ── LEASE RENEWAL (EOM logic) ─────────────────────────────────────────────────
lr_rows  = read_sheet(FILES['lr'])
lr_eom   = eom_we_map(lr_rows, 14, 15, 16)
lr_r_tbl = defaultdict(int); lr_p_tbl = defaultdict(int); lr_m_tbl = defaultdict(int)
for r in lr_rows[1:]:
    if not r or not r[14] or not r[15] or not r[16]: continue
    year  = str(r[14]).strip(); month = str(r[15]).strip(); we = r[16]
    if lr_eom.get((year,month)) != we: continue
    attrs  = resolve(r[17] or r[0]) or {}
    ter    = str(r[2]).strip() if r[2] else attrs.get('territory','')
    ter    = {'NORTH OC':'North OC','SOUTH OC':'South OC','SD PROPERTIES':'SD Properties',
              'COMMERCIAL':'Commercial','BRENDEN':'Brenden','ELDERKIN':'Elderkin'}.get(ter.upper(), ter)
    prop   = attrs.get('proptype','')
    status = str(r[10]).strip() if r[10] else ''
    tbl    = lr_r_tbl if status=='Renewed' else lr_p_tbl if status=='Pending' else              lr_m_tbl if status in ('Move-Out','Changed to MTM') else None
    if tbl is None: continue
    for k in agg_keys(year,month,ter,prop): tbl[k] += 1

DATA['lr_r'] = dict(lr_r_tbl); DATA['lr_p'] = dict(lr_p_tbl); DATA['lr_m'] = dict(lr_m_tbl)

# WE → month map (for WE filter on LR KPI)
we_month_map = {}
for (yr,mo), eom in lr_eom.items():
    we_month_map[eom.strftime('%Y-%m-%d')] = {'year': yr, 'month': mo}
DATA['we_month'] = we_month_map

# ── UNIT TURN (EOM: blank comments + posted to internet) ─────────────────────
ut_rows = read_sheet(FILES['unit_turn'], 'Unit Turn Details')
ut_tbl  = defaultdict(lambda: {'sum':0,'cnt':0})
for r in ut_rows[1:]:
    if not r or not r[17]: continue
    if r[16] is not None and str(r[16]).strip() != '': continue  # skip if comment
    if r[25] is None: continue                                      # skip if not posted
    attrs = resolve(r[0]) or {}
    ter   = attrs.get('territory',''); prop = attrs.get('proptype','')
    days  = r[17] if isinstance(r[17],(int,float)) else 0
    month = str(r[20]).strip() if r[20] else ''
    year  = str(r[19]).strip() if r[19] else ''
    for k in agg_keys(year,month,ter,prop):
        ut_tbl[k]['sum'] += days; ut_tbl[k]['cnt'] += 1
DATA['ut'] = {k: round(v['sum']/v['cnt'],2) for k,v in ut_tbl.items() if v['cnt']>0}

# ── NPS (from New_NPS_Data, formula: (P-D)/T*100) ─────────────────────────────
nps_rows = read_sheet(FILES['new_nps'])
nps_raw  = defaultdict(lambda: {'P':0,'D':0,'T':0})
for r in nps_rows[1:]:
    if not r or not r[11] or not r[37]: continue
    ter   = str(r[11]).strip(); month = str(r[37]).strip()
    year  = str(r[36]).strip() if r[36] else ''
    p = 1 if r[39]==1 else 0; pa = 1 if r[40]==1 else 0; d = 1 if r[41]==1 else 0
    t = p+pa+d
    if t==0: continue
    for k in agg_keys(year,month,ter,''):
        nps_raw[k]['P']+=p; nps_raw[k]['D']+=d; nps_raw[k]['T']+=t
DATA['nps'] = {k: round((v['P']-v['D'])/v['T']*100,1) for k,v in nps_raw.items() if v['T']>0}

print(f"  DATA tables: {list(DATA.keys())}")

# ── SECTION 4: BUILD D2 SUPPLEMENTARY TABLES ─────────────────────────────────
print("[4/8] Computing D2 tables (vacancy, apps, WO status, reviews, BDM, unpaid, AR%)...")

D2 = {}

# ── VACANCY DONUT (EOM + WE snapshots from Vacancy Rate sheet) ───────────────
vac_rate  = read_sheet(FILES['vacancy'], 'Vacancy Rate')
ter_map1  = read_sheet(FILES['vacancy'], 'Territory Mapping1')
addr_ter  = {str(r[0]).strip(): str(r[1]).strip() for r in ter_map1[1:] if r and r[0] and r[1]}
vac_eom   = eom_we_map(vac_rate, 17, 16, 18)

def get_vac_ter(r):
    ter = str(r[25]).strip() if r[25] else ''
    if ter and ter != 'None': return ter
    addr = str(r[15]).strip() if r[15] else str(r[0]).strip()
    return addr_ter.get(addr, '') or (resolve(addr) or {}).get('territory','')

vac_donut2 = defaultdict(lambda: defaultdict(int))
all_vac_we = sorted(set(r[18] for r in vac_rate[1:] if r and r[18] and hasattr(r[18],'strftime')))
latest_vac_we = all_vac_we[-1]

for r in vac_rate[1:]:
    if not r or not r[18] or not hasattr(r[18],'strftime'): continue
    if str(r[2]).strip() != 'Vacant-Unrented': continue
    year  = str(r[17]).strip(); month = str(r[16]).strip()
    we    = r[18]; we_str = we.strftime('%Y-%m-%d')
    ter   = get_vac_ter(r)
    if ter in SKIP_TERS: continue
    vac_donut2[f"we:{we_str}"][ter] += 1
    if vac_eom.get((year,month)) == we:
        vac_donut2[f"eom:{year}|{month}"][ter] += 1
        vac_donut2[f"eom:{year}|"][ter]         += 1
        vac_donut2[f"eom:|{month}"][ter]         += 1
        vac_donut2["eom:|||"][ter]               += 1

D2['vac_donut2'] = {k: dict(v) for k,v in vac_donut2.items()}
D2['latest_we']  = latest_vac_we.strftime('%Y-%m-%d')

# ── PORTFOLIO COUNTS (row count per territory per WE — denominator for vacancy %) ──
port_eom = eom_we_map(port_rows, 12, 11, 13)
port_cnt  = defaultdict(lambda: defaultdict(int))
for r in port_rows[1:]:
    if not r or not r[6] or not r[13] or not hasattr(r[13],'strftime'): continue
    ter    = str(r[6]).strip()
    if ter in SKIP_TERS: continue
    year   = str(r[12]).strip(); month = str(r[11]).strip()
    we     = r[13]; we_str = we.strftime('%Y-%m-%d')
    port_cnt[f"we:{we_str}"][ter] += 1
    if port_eom.get((year,month)) == we:
        port_cnt[f"eom:{year}|{month}"][ter] += 1
        port_cnt[f"eom:{year}|"][ter]         += 1
        port_cnt[f"eom:|{month}"][ter]         += 1
        port_cnt["eom:|||"][ter]               += 1

D2['port_counts'] = {k: dict(v) for k,v in port_cnt.items()}

# ── VAC_TOP (latest WE snapshot for Top Vacant table) ────────────────────────
def make_display(prop, unit):
    p = str(prop).strip() if prop else ''
    u = str(unit).strip() if unit else ''
    return f"{p} - {u}" if u and u != p else p

vac_top = []
vac_we_snaps = defaultdict(list)
for r in vac_rate[1:]:
    if not r or not r[18] or not hasattr(r[18],'strftime'): continue
    if str(r[2]).strip() != 'Vacant-Unrented': continue
    we_str = r[18].strftime('%Y-%m-%d')
    ter    = get_vac_ter(r)
    attrs  = resolve(r[0]) or {}
    entry  = {
        'name': make_display(r[0], r[1]),
        'days': int(r[4]) if isinstance(r[4],(int,float)) and r[4] else 0,
        'territory': ter, 'proptype': attrs.get('proptype',''),
        'month': str(r[16]).strip() if r[16] else '',
        'year':  str(r[17]).strip() if r[17] else '',
        'status': 'Vacant-Unrented',
    }
    if r[18] == latest_vac_we:
        vac_top.append(entry)
    snap = {k: v for k,v in entry.items() if k not in ('month','year','status')}
    vac_we_snaps[we_str].append(snap)

vac_top.sort(key=lambda x: -x['days'])
for k in vac_we_snaps: vac_we_snaps[k].sort(key=lambda x: -x['days'])
D2['vac_top'] = vac_top
D2['vac_we']  = dict(vac_we_snaps)

# ── APPLICATIONS ──────────────────────────────────────────────────────────────
app_rows = read_sheet(FILES['applications'])
app_tbl  = defaultdict(lambda: [0,0,0,0])  # [canceled,denied,approved,converted]
STATUS_IDX = {'Canceled':0, 'Denied':1, 'Approved':2, 'Converted':3}
for r in app_rows[1:]:
    if not r: continue
    addr  = r[27] or r[0]
    attrs = resolve(str(addr)) if addr else {}
    if not attrs: attrs = {}
    status = str(r[8]).strip().title() if r[8] else ''
    ter    = attrs.get('territory',''); prop = attrs.get('proptype','')
    year   = str(r[30]).strip() if r[30] else ''
    month  = str(r[29]).strip() if r[29] else ''
    idx = STATUS_IDX.get(status, -1)
    if idx < 0: continue
    for k in agg_keys(year,month,ter,prop): app_tbl[k][idx] += 1
D2['app_donut'] = {k: list(v) for k,v in app_tbl.items()}

# ── WORK ORDER STATUS + ISSUES ────────────────────────────────────────────────
wo_status_tbl = defaultdict(lambda: defaultdict(int))
wo_issue_tbl  = defaultdict(lambda: defaultdict(int))
for r in wo_rows[1:]:
    if not r or not r[0]: continue
    attrs  = resolve(r[0]) or {}
    ter    = attrs.get('territory',''); prop = attrs.get('proptype','')
    year   = str(r[42]).strip() if r[42] else ''
    month  = str(r[43]).strip() if r[43] else ''
    status = str(r[10]).strip() if r[10] else ''
    issue  = str(r[29]).strip() if r[29] else ''
    for k in agg_keys(year,month,ter,prop):
        if status: wo_status_tbl[k][status] += 1
        if issue:  wo_issue_tbl[k][issue]   += 1
D2['wo_status'] = {k: dict(v) for k,v in wo_status_tbl.items()}
D2['wo_issues'] = {k: dict(sorted(v.items(),key=lambda x:-x[1])[:10]) for k,v in wo_issue_tbl.items()}

# ── CONCESSIONS ───────────────────────────────────────────────────────────────
con_rows = read_sheet(FILES['concessions'])
D2['concessions'] = []
for r in con_rows[1:]:
    if not r or not r[0]: continue
    attrs = resolve(r[7] or r[0]) or {}
    D2['concessions'].append({
        'property': str(r[0]).strip(),
        'amount':   r[1] if isinstance(r[1],(int,float)) else 0,
        'month':    str(r[3]).strip() if r[3] else '',
        'year':     str(r[5]).strip() if r[5] else '',
        'we':       r[4].strftime('%Y-%m-%d') if r[4] and hasattr(r[4],'strftime') else '',
        'reason':   str(r[6]).strip() if r[6] else '',
        'territory':attrs.get('territory',''),
        'proptype': attrs.get('proptype',''),
    })

# ── UNPAID CHARGES (EOM logic) ────────────────────────────────────────────────
up_rows = read_sheet(FILES['unpaid'])
up_eom  = eom_we_map(up_rows, 15, 14, 16)
up_flat = []; up_totals = defaultdict(lambda: {'total':0.0,'count':0})
seen_up = set()
for r in up_rows[1:]:
    if not r or not r[16] or not hasattr(r[16],'strftime'): continue
    year  = str(r[15]).strip(); month = str(r[14]).strip(); we = r[16]
    if up_eom.get((year,month)) != we: continue
    amt   = r[8] if isinstance(r[8],(int,float)) and r[8] and r[8]>0 else 0
    if amt == 0: continue
    attrs = resolve(r[1] or r[0]) or {}
    ter   = str(r[17]).strip() if r[17] else attrs.get('territory','')
    prop  = attrs.get('proptype','')
    entry = {
        'n': str(r[4]).strip() if r[4] else '',
        'pr':str(r[0]).strip() if r[0] else '',
        'u': str(r[3]).strip() if r[3] else '',
        'a': round(amt,2), 's': str(r[5]).strip() if r[5] else '',
        't': ter, 'pt': prop, 'mo': month, 'yr': year,
        'we': we.strftime('%Y-%m-%d'),
    }
    up_flat.append(entry)
    for k in agg_keys(year,month,ter,prop):
        up_totals[k]['total'] += amt; up_totals[k]['count'] += 1
up_flat.sort(key=lambda x: -x['a'])
D2['unpaid']   = up_flat
D2['up_totals']= {k:{'total':round(v['total'],2),'count':v['count']} for k,v in up_totals.items()}

# ── RENT ROLL (avg rent per unit) ─────────────────────────────────────────────
rr_rows  = read_sheet(FILES['rent_roll'])
rr_rent  = defaultdict(float); rr_count = defaultdict(int)
for r in rr_rows[1:]:
    if not r or not r[0]: continue
    ter   = str(r[22]).strip() if r[22] else ''
    year  = str(r[20]).strip() if r[20] else ''
    month = str(r[19]).strip() if r[19] else ''
    rent  = r[9] if isinstance(r[9],(int,float)) and r[9] and r[9]>0 else 0
    for k in [f"{year}|{month}|{ter}|",f"{year}|{month}||",
              f"|{month}|{ter}|",f"|{month}||",
              f"{year}||{ter}|",f"{year}|||",f"||{ter}|","|||"]:
        rr_rent[k] += rent; rr_count[k] += 1
D2['rr_rent']  = {k: round(v,2) for k,v in rr_rent.items()}
D2['rr_count'] = dict(rr_count)
D2['rr_avg']   = {k: round(rr_rent[k]/rr_count[k],2) for k in rr_count if rr_count[k]>0}

# ── AR % ──────────────────────────────────────────────────────────────────────
ar_pct = {}
for k in up_totals:
    ar  = up_totals[k]['total']
    rnt = rr_rent.get(k, 0)
    if rnt > 0: ar_pct[k] = round(ar/rnt*100, 2)
D2['ar_pct'] = ar_pct

# ── 5-STAR REVIEWS ────────────────────────────────────────────────────────────
star_rows = read_sheet(FILES['five_star'], 'rawdata reviews')
staff_ter = defaultdict(set)
name_norm = {'Anthony Ruiz':'Anthony Andrew Ruiz'}
for r in port_rows[1:]:
    if not r or not r[6]: continue
    ter = str(r[6]).strip()
    for col in [7,8,9,10]:
        if r[col] and str(r[col]).strip() not in ('TBD','TBD_Leasing Agent','Commercial','STONE PM',''):
            staff_ter[str(r[col]).strip()].add(ter)

def norm_plat(p):
    p = str(p).strip().lower()
    if 'yelp' in p and 'oc' in p: return 'Yelp OC'
    if 'google' in p and 'oc' in p: return 'Google OC'
    if 'google' in p and 'sd' in p: return 'Google SD'
    if 'yelp' in p and 'sd' in p: return 'Yelp SD'
    return 'Other'

rev_agg = defaultdict(lambda: defaultdict(lambda: {'c':0,'r':0,'yo':0,'go':0,'gs':0,'ys':0}))
plat_agg= defaultdict(lambda: defaultdict(list))
for r in star_rows[1:]:
    if not r or not r[0]: continue
    name   = str(r[0]).strip()
    if name == 'TBD': continue
    mapped = name_norm.get(name,name)
    ters   = sorted(staff_ter.get(mapped, {''}))
    plat   = norm_plat(r[4]) if r[4] else 'Other'
    rating = r[5] if isinstance(r[5],(int,float)) else 5
    year   = str(r[3]).strip() if r[3] else ''
    month  = str(r[2]).strip() if r[2] else ''
    pk     = {'Yelp OC':'yo','Google OC':'go','Google SD':'gs','Yelp SD':'ys'}.get(plat,'o')
    ter    = ters[0] if ters and ters[0] else ''
    for k in agg_keys(year,month,ter,''):
        rev_agg[k][name]['c']  += 1
        rev_agg[k][name]['r']  += rating
        rev_agg[k][name][pk]   += 1
        plat_agg[k][plat].append(rating)

D2['review_lb'] = {
    k: sorted([{'n':n,'c':s['c'],'r':round(s['r']/s['c'],2) if s['c'] else 5,
                'yo':s['yo'],'go':s['go'],'gs':s['gs'],'ys':s['ys']}
               for n,s in names.items()], key=lambda x:-x['c'])[:15]
    for k,names in rev_agg.items()
}
D2['plat_ratings'] = {
    k: {p: round(sum(v)/len(v),2) for p,v in pd.items()}
    for k,pd in plat_agg.items()
}

# ── BDM ───────────────────────────────────────────────────────────────────────
bdm_rows   = read_sheet(FILES['bdm'])
bdm_status = defaultdict(lambda: defaultdict(int))
bdm_source = defaultdict(lambda: defaultdict(int))
bdm_all    = []
for r in bdm_rows[1:]:
    if not r or not r[0]: continue
    status = str(r[3]).strip().lower() if r[3] else ''
    source = str(r[4]).strip()         if r[4] else ''
    date_val = next((r[i] for i in range(len(r)) if r[i] and hasattr(r[i],'strftime')), None)
    month  = date_val.strftime('%B') if date_val else ''
    year   = date_val.strftime('%Y') if date_val else ''
    for k in agg_keys(year,month,'',''):
        bdm_status[k][status] += 1
        bdm_source[k][source] += 1
    bdm_all.append({'status':status,'source':source,'month':month,'year':year,'property':str(r[0]).strip()})
D2['bdm_status'] = {k: dict(v) for k,v in bdm_status.items()}
D2['bdm_source'] = {k: dict(v) for k,v in bdm_source.items()}
D2['bdm_all']    = bdm_all

print(f"  D2 tables: {list(D2.keys())}")

# ── SECTION 5: COMPUTE SUMMARY STATS FOR KPI DEFAULTS ────────────────────────
print("[5/8] Computing summary stats...")
build_date = datetime.datetime.now().strftime('%Y-%m-%d %H:%M UTC')

# ── SECTION 6: ASSEMBLE HTML ──────────────────────────────────────────────────
print("[6/8] Assembling dashboard HTML...")

# Use placeholder replacement instead of f-string to avoid conflicts
# with CSS and JS curly braces
HTML_TEMPLATE = (
    '<!DOCTYPE html>\n'
    '<html lang="en">\n'
    '<head>\n'
    '<meta charset="UTF-8">\n'
    '<meta name="viewport" content="width=device-width, initial-scale=1.0">\n'
    '<title>AllView Real Estate — Dashboard</title>\n'
    '<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.1/dist/chart.umd.min.js"></script>\n'
    '<style><<<CSS>>></style>\n'
    '</head>\n'
    '<body>\n'
    '<<<BODY>>>\n'
    '<script>\n'
    '<<<DATA>>>\n'
    '<<<D2>>>\n'
    '<<<JS>>>\n'
    '</script>\n'
    '</body>\n'
    '</html>'
)

CSS = """ + repr(css) + """

BODY_HTML = """ + repr(body_html) + """

JS_LOGIC = """ + repr(js_logic) + """

dashboard = (HTML_TEMPLATE
    .replace('<<<CSS>>>',  CSS)
    .replace('<<<BODY>>>', BODY_HTML)
    .replace('<<<DATA>>>', 'const DATA = ' + json.dumps(DATA) + ';\n')
    .replace('<<<D2>>>',   'const D2 = '   + json.dumps(D2)   + ';\n')
    .replace('<<<JS>>>',   JS_LOGIC)
)

# ── SECTION 7: WRITE OUTPUT ───────────────────────────────────────────────────
print("[7/8] Writing output file...")
out_path = os.path.join(OUTPUT_DIR, 'index.html')
with open(out_path, 'w', encoding='utf-8') as f:
    f.write(dashboard)

size_kb = len(dashboard) / 1024
print(f"  Written: {out_path} ({size_kb:.0f} KB)")

# ── SECTION 8: DONE ───────────────────────────────────────────────────────────
print("[8/8] Done!")
print('=' * 60)
print(f'Dashboard built: {build_date}')
print(f'Output:          {out_path}')
print(f'Size:            {size_kb:.0f} KB')
print('=' * 60)
