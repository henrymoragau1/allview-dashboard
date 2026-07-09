"""
build.py - AllView Real Estate Dashboard Builder
Reads Excel files from ./data/, outputs ./output/index.html
"""

import os, re, sys, json, datetime, datetime
import openpyxl
from collections import defaultdict

DATA_DIR     = './data'
OUTPUT_DIR   = './output'
TEMPLATE_PATH = './scripts/dashboard_template.html'
os.makedirs(OUTPUT_DIR, exist_ok=True)

print("=" * 60)
print("AllView Dashboard Builder")
print("=" * 60)

FILES = {
    'portfolio':    os.path.join(DATA_DIR, 'New_Portafolios_Structure_-_Effective_Feb_17.xlsx'),
    'vacancy':      os.path.join(DATA_DIR, 'Rental_Vacancy.xlsx'),
    'applications': os.path.join(DATA_DIR, 'rental_applications-Performance.xlsx'),
    'move_inout':   os.path.join(DATA_DIR, 'tenant_tickler_-_Move-in-out.xlsx'),
    'unpaid':       os.path.join(DATA_DIR, 'tenant_unpaid_charges_summary.xlsx'),
    'unit_turn':    os.path.join(DATA_DIR, 'Unit_Turn_Details.xlsx'),
    'work_orders':  os.path.join(DATA_DIR, 'WorkOrders_Report.xlsx'),
    'five_star':    os.path.join(DATA_DIR, '5-star_Tracker.xlsx'),
    'bdm':          os.path.join(DATA_DIR, 'BDM_Data.xlsx'),
    'lr':           os.path.join(DATA_DIR, 'LR_Performance.xlsx'),
    'new_nps':      os.path.join(DATA_DIR, 'New_NPS_Data.xlsx'),
    'concessions':  os.path.join(DATA_DIR, 'Owner_Concessions.xlsx'),
    'rent_roll':    os.path.join(DATA_DIR, 'Rent_Roll.xlsx'),
    'phone':        os.path.join(DATA_DIR, 'Phone_Reporting.xlsx'),
    'term':         os.path.join(DATA_DIR, 'property_directory-Term.xlsx'),
    'guest_card':   os.path.join(DATA_DIR, 'guest_card_inquiries-Performance.xlsx'),
    'sm_active':    os.path.join(DATA_DIR, 'Active_Listings_-_Performance.xlsx'),
    'sm_offmkt':    os.path.join(DATA_DIR, 'Off_Market_Listings_-_Performance__Historical_Data_.xlsx'),
}

missing = [k for k,v in FILES.items() if not os.path.exists(v)]
if missing: print(f"WARNING: Missing: {missing}")
else:       print(f"[1/7] All source files found")

def read_sheet(path, sheet=0):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb[sheet] if isinstance(sheet, str) else wb.worksheets[sheet]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    return rows

def eom_we_map(rows, year_col, month_col, we_col):
    # EOM WE = last WE in each calendar month, EXCLUDING WEs on the last day of the month
    # (WEs on month's last day are treated as first week of next month in reporting)
    import calendar as cal_mod
    MNAMES={1:'January',2:'February',3:'March',4:'April',5:'May',6:'June',
            7:'July',8:'August',9:'September',10:'October',11:'November',12:'December'}
    all_wes=list(set(r[we_col] for r in rows if r and r[we_col] and hasattr(r[we_col],'strftime')))
    eom={}
    for yr in range(2024,2028):
        for mo_num in range(1,13):
            last_day=cal_mod.monthrange(yr,mo_num)[1]
            month_wes=[w for w in all_wes if w.year==yr and w.month==mo_num and w.day<last_day]
            if month_wes: eom[(str(yr),MNAMES[mo_num])]=max(month_wes)
    return eom

def get_eom_key(we_dt, eom):
    import calendar as cal_mod
    MNAMES={1:'January',2:'February',3:'March',4:'April',5:'May',6:'June',
            7:'July',8:'August',9:'September',10:'October',11:'November',12:'December'}
    yr=str(we_dt.year); mo=MNAMES[we_dt.month]
    return (yr,mo) if eom.get((yr,mo))==we_dt else None

def agg_keys(year, month, ter, prop):
    y,m,t,p = year,month,ter,prop
    return [
        f"{y}|{m}|{t}|{p}", f"{y}|{m}|{t}|", f"{y}|{m}||{p}", f"{y}|{m}||",
        f"|{m}|{t}|{p}",    f"|{m}|{t}|",    f"|{m}||{p}",    f"|{m}||",
        f"{y}||{t}|{p}",    f"{y}||{t}|",    f"{y}|||{p}",    f"{y}|||",
        f"||{t}|{p}",       f"||{t}|",       f"|||{p}",       "|||",
    ]

MONTHS    = ['January','February','March','April','May','June',
             'July','August','September','October','November','December']
SKIP_TERS = {'TBD','Mid OC',''}

print("[2/7] Building portfolio master mapping...")
port_rows    = read_sheet(FILES['portfolio'])
master_full  = {}
master_short = defaultdict(list)
for r in port_rows[1:]:
    if not r or not r[0]: continue
    addr  = str(r[0]).strip()
    attrs = {'territory':str(r[6]).strip() if r[6] else '',
             'proptype': str(r[5]).strip() if r[5] else '',
             'pm':       str(r[8]).strip() if r[8] else '',
             'la':       str(r[7]).strip() if r[7] else ''}
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
    base  = re.split(r'\s*[\|#]\s*', name)[0].strip()
    base  = re.sub(r'\s+(unit|apt|suite|ste)\s+\S+$','',base,flags=re.I).strip()
    short2 = base.split(',')[0].strip().lower()
    if short2 in master_short: return master_full[master_short[short2][0]]
    if ' - ' in name:
        fp = name.split(' - ',1)[1].strip()
        if fp in master_full: return master_full[fp]
        s3 = fp.split(',')[0].strip().lower()
        if s3 in master_short: return master_full[master_short[s3][0]]
    return None

print("[3/7] Computing DATA tables...")
DATA = {}

# Move In/Out
# Count based on col X (Month) only — no year filter per user spec
# Each row in col D (Event) = one move-in or move-out regardless of year
mio_rows = read_sheet(FILES['move_inout'])
mi_tbl=defaultdict(int); mo_tbl=defaultdict(int)
mi_we=defaultdict(int);  mo_we=defaultdict(int)
for r in mio_rows[1:]:
    if not r or not r[0] or len(r) < 24: continue
    attrs=resolve(r[0]) or {}
    ter=attrs.get('territory',''); prop=attrs.get('proptype','')
    event=str(r[3]).strip() if r[3] else ''
    month=str(r[23]).strip() if r[23] else ''   # col X = Month (no year filter)
    year=str(r[22]).strip() if r[22] else ''     # col W = Year (kept for WE filter)
    we=r[24].strftime('%Y-%m-%d') if r[24] and hasattr(r[24],'strftime') else ''
    if event not in ('Move-in','Move-out'): continue
    tbl=mi_tbl if event=='Move-in' else mo_tbl
    # Month-only keys (no year) — matches user's manual count
    for k in [f"|{month}|{ter}|{prop}", f"|{month}|{ter}|",
              f"|{month}||{prop}",       f"|{month}||"]:
        tbl[k] += 1
    # Year+month keys for year filter
    if year:
        for k in [f"{year}|{month}|{ter}|{prop}", f"{year}|{month}|{ter}|",
                  f"{year}|{month}||{prop}",       f"{year}|{month}||",
                  f"{year}||{ter}|{prop}",          f"{year}||{ter}|",
                  f"{year}|||{prop}",               f"{year}|||"]:
            tbl[k] += 1
    # All-time rollup
    for k in [f"||{ter}|{prop}", f"||{ter}|", f"|||{prop}", "|||"]:
        tbl[k] += 1
    # WE level
    wt=mi_we if event=='Move-in' else mo_we
    if we:
        wt[f"{we}|{ter}|{prop}"]+=1; wt[f"{we}|{ter}|"]+=1; wt[f"{we}||"]+=1
DATA['mi']=dict(mi_tbl); DATA['mo']=dict(mo_tbl)
DATA['mi_we']=dict(mi_we); DATA['mo_we']=dict(mo_we)

# Work Orders — Created, Closed, Opened
# Created = col 14 Created At date
# Closed  = col 24 Completed On date
# Opened  = col 48 Opened WO = 1, grouped by col 42/43 Year/Month
wo_rows=read_sheet(FILES['work_orders'])
wo_created=defaultdict(int); wo_closed=defaultdict(int)
wo_opened=defaultdict(int);  wo_we_tbl=defaultdict(int)

for r in wo_rows[1:]:
    if not r or not r[0]: continue
    attrs=resolve(r[0]) or {}
    ter=attrs.get('territory',''); prop=attrs.get('proptype','')
    we=r[44].strftime('%Y-%m-%d') if len(r)>44 and r[44] and hasattr(r[44],'strftime') else ''

    # CREATED — by Created At date (col 14)
    if r[14] and hasattr(r[14],'strftime'):
        yr=r[14].strftime('%Y'); mo=r[14].strftime('%B')
        for k in agg_keys(yr,mo,ter,prop): wo_created[k]+=1

    # CLOSED — by Completed On date (col 24)
    if r[24] and hasattr(r[24],'strftime'):
        yr=r[24].strftime('%Y'); mo=r[24].strftime('%B')
        for k in agg_keys(yr,mo,ter,prop): wo_closed[k]+=1

    # OPENED — col 48 = 1, by Year/Month cols
    if len(r)>48 and r[48] and str(r[48]).strip()=='1':
        yr=str(r[42]).strip() if r[42] else ''
        mo=str(r[43]).strip() if r[43] else ''
        if yr and mo:
            for k in agg_keys(yr,mo,ter,prop): wo_opened[k]+=1

    # WE level (created date)
    if we:
        wo_we_tbl[f"{we}|{ter}|{prop}"]+=1
        wo_we_tbl[f"{we}|{ter}|"]+=1
        wo_we_tbl[f"{we}||"]+=1

DATA['wo']=dict(wo_created)   # keep 'wo' key for WO KPI total
DATA['wo_created']=dict(wo_created)
DATA['wo_closed']=dict(wo_closed)
DATA['wo_opened']=dict(wo_opened)
DATA['wo_we']=dict(wo_we_tbl)

# Lease Renewal (EOM)
lr_rows=read_sheet(FILES['lr'])
lr_eom=eom_we_map(lr_rows,14,15,16)
lr_r=defaultdict(int); lr_p=defaultdict(int); lr_m=defaultdict(int)
for r in lr_rows[1:]:
    if not r or not r[14] or not r[15] or not r[16]: continue
    year=str(r[14]).strip(); month=str(r[15]).strip(); we=r[16]
    if lr_eom.get((year,month))!=we: continue
    attrs=resolve(r[17] or r[0]) or {}
    ter=str(r[2]).strip() if r[2] else attrs.get('territory','')
    ter={'NORTH OC':'North OC','SOUTH OC':'South OC','SD PROPERTIES':'SD Properties',
         'COMMERCIAL':'Commercial','BRENDEN':'Brenden','ELDERKIN':'Elderkin'}.get(ter.upper(),ter)
    prop=attrs.get('proptype','')
    status=str(r[10]).strip() if r[10] else ''
    tbl=lr_r if status=='Renewed' else lr_p if status=='Pending' else \
        lr_m if status in ('Move-Out','Changed to MTM') else None
    if tbl is None: continue
    for k in agg_keys(year,month,ter,prop): tbl[k]+=1
DATA['lr_r']=dict(lr_r); DATA['lr_p']=dict(lr_p); DATA['lr_m']=dict(lr_m)
we_month_map={}
for (yr,mo),eom in lr_eom.items():
    we_month_map[eom.strftime('%Y-%m-%d')]={'year':yr,'month':mo}
DATA['we_month']=we_month_map

# Unit Turn
ut_rows=read_sheet(FILES['unit_turn'],'Unit Turn Details')
ut_tbl=defaultdict(lambda:{'sum':0,'cnt':0})
for r in ut_rows[1:]:
    if not r or not r[17]: continue
    if r[16] is not None and str(r[16]).strip()!='': continue  # blank comment
    if not r[25]: continue  # posted to internet at must have a date
    attrs=resolve(r[0]) or {}
    ter=attrs.get('territory',''); prop=attrs.get('proptype','')
    days=r[17] if isinstance(r[17],(int,float)) else 0
    month=str(r[20]).strip() if r[20] else ''
    year=str(r[19]).strip() if r[19] else ''
    for k in agg_keys(year,month,ter,prop):
        ut_tbl[k]['sum']+=days; ut_tbl[k]['cnt']+=1
DATA['ut']={k:round(v['sum']/v['cnt'],2) for k,v in ut_tbl.items() if v['cnt']>0}

# NPS
nps_rows=read_sheet(FILES['new_nps'])
nps_raw=defaultdict(lambda:{'P':0,'D':0,'T':0})
for r in nps_rows[1:]:
    if not r or not r[11] or not r[37]: continue
    ter=str(r[11]).strip(); month=str(r[37]).strip()
    year=str(r[36]).strip() if r[36] else ''
    p=1 if r[39]==1 else 0; pa=1 if r[40]==1 else 0; d=1 if r[41]==1 else 0
    t=p+pa+d
    if t==0: continue
    for k in agg_keys(year,month,ter,''):
        nps_raw[k]['P']+=p; nps_raw[k]['D']+=d; nps_raw[k]['T']+=t
DATA['nps']={k:round((v['P']-v['D'])/v['T']*100,1) for k,v in nps_raw.items() if v['T']>0}

# ── NPS CUMULATIVE BY TERRITORY PER WE ───────────────────────────────────────
# Used by CEO Scorecard for all-period cumulative NPS
nps_rows2=read_sheet(FILES['new_nps'])
nps_we_entries=defaultdict(list)
for r in nps_rows2[1:]:
    if not r or not r[35] or not hasattr(r[35],'strftime'): continue
    p=1 if r[39]==1 else 0; d=1 if r[41]==1 else 0; pa=1 if r[40]==1 else 0
    t=p+pa+d
    if t==0: continue
    ter=str(r[11]).strip() if r[11] else ''
    nps_we_entries[r[35].strftime('%Y-%m-%d')].append({'ter':ter,'p':p,'d':d,'t':t})

nps_all_wes=sorted(nps_we_entries.keys())
TERS_NPS=['North OC','South OC','SD Properties','Commercial','Brenden','Elderkin','STONE']
ter_run=defaultdict(lambda:{'P':0,'D':0,'T':0})
all_run={'P':0,'D':0,'T':0}
nps_cum_ter={ter:{} for ter in TERS_NPS}
nps_cum_all={}
for we in nps_all_wes:
    for e in nps_we_entries[we]:
        ter=e['ter']
        all_run['P']+=e['p']; all_run['D']+=e['d']; all_run['T']+=e['t']
        if ter in SKIP_TERS: continue
        ter_run[ter]['P']+=e['p']; ter_run[ter]['D']+=e['d']; ter_run[ter]['T']+=e['t']
    for ter in TERS_NPS:
        v=ter_run[ter]
        nps_cum_ter[ter][we]=round((v['P']-v['D'])/v['T']*100,1) if v['T']>0 else None
    nps_cum_all[we]=round((all_run['P']-all_run['D'])/all_run['T']*100,1) if all_run['T']>0 else None
DATA['nps_cum_ter']=nps_cum_ter
DATA['nps_cum_all']=nps_cum_all
print(f"  DATA: {list(DATA.keys())}")

print("[4/7] Computing D2 tables...")
D2={}

# Vacancy
vac_rate=read_sheet(FILES['vacancy'],'Vacancy Rate')
ter_map1=read_sheet(FILES['vacancy'],'Territory Mapping1')
addr_ter={str(r[0]).strip():str(r[1]).strip() for r in ter_map1[1:] if r and r[0] and r[1]}
vac_eom=eom_we_map(vac_rate,17,16,18)

def get_vac_ter(r):
    # Portfolio master is authoritative — addr_ter mapping1 may have stale Mid OC
    addr=str(r[15]).strip() if r[15] else str(r[0]).strip()
    resolved=(resolve(addr) or {}).get('territory','')
    if resolved and resolved not in SKIP_TERS: return resolved
    ter=str(r[25]).strip() if r[25] else ''
    if ter and ter not in SKIP_TERS: return ter
    mapped=addr_ter.get(addr,'')
    return mapped if mapped not in SKIP_TERS else ''

vac_donut2=defaultdict(lambda:defaultdict(int))
all_vac_we=sorted(set(r[18] for r in vac_rate[1:] if r and r[18] and hasattr(r[18],'strftime')))
latest_vac_we=all_vac_we[-1]
for r in vac_rate[1:]:
    if not r or not r[18] or not hasattr(r[18],'strftime'): continue
    if str(r[2]).strip()!='Vacant-Unrented': continue
    year=str(r[17]).strip(); month=str(r[16]).strip()
    we=r[18]; we_str=we.strftime('%Y-%m-%d')
    ter=get_vac_ter(r)
    if ter in SKIP_TERS: continue
    vac_donut2[f"we:{we_str}"][ter]+=1
    if vac_eom.get((year,month))==we:
        vac_donut2[f"eom:{year}|{month}"][ter]+=1
        vac_donut2[f"eom:{year}|"][ter]+=1
        vac_donut2[f"eom:|{month}"][ter]+=1
        vac_donut2["eom:|||"][ter]+=1
D2['vac_donut2']={k:dict(v) for k,v in vac_donut2.items()}
# Validate latest_we: must have full portfolio data (>100 units in port_counts)
_latest_we_str=latest_vac_we.strftime('%Y-%m-%d')
# If the latest vacancy WE has no matching portfolio data, step back to prior WE
_valid_we_keys=[k.replace('we:','') for k,v in D2.get('port_counts',{}).items()
                if k.startswith('we:') and sum(v.values())>100]
if _latest_we_str not in _valid_we_keys and _valid_we_keys:
    _latest_we_str=sorted(_valid_we_keys)[-1]
D2['latest_we']=_latest_we_str

# Portfolio counts
port_eom=eom_we_map(port_rows,12,11,13)
port_cnt=defaultdict(lambda:defaultdict(int))
for r in port_rows[1:]:
    if not r or not r[6] or not r[13] or not hasattr(r[13],'strftime'): continue
    ter=str(r[6]).strip()
    if ter in SKIP_TERS: continue
    year=str(r[12]).strip(); month=str(r[11]).strip()
    we=r[13]; we_str=we.strftime('%Y-%m-%d')
    port_cnt[f"we:{we_str}"][ter]+=1
    if port_eom.get((year,month))==we:
        port_cnt[f"eom:{year}|{month}"][ter]+=1
        port_cnt[f"eom:{year}|"][ter]+=1
        port_cnt[f"eom:|{month}"][ter]+=1
        port_cnt["eom:|||"][ter]+=1
# Remove 'we:' keys with fewer than 100 units (partial/bad snapshots)
D2['port_counts']={k:dict(v) for k,v in port_cnt.items()
                   if not k.startswith('we:') or sum(v.values())>=100}

# Vac top
def make_display(prop,unit):
    p=str(prop).strip() if prop else ''; u=str(unit).strip() if unit else ''
    return f"{p} - {u}" if u and u!=p else p

vac_top=[]; vac_we_snaps=defaultdict(list)
for r in vac_rate[1:]:
    if not r or not r[18] or not hasattr(r[18],'strftime'): continue
    if str(r[2]).strip()!='Vacant-Unrented': continue
    we_str=r[18].strftime('%Y-%m-%d')
    ter=get_vac_ter(r); attrs=resolve(r[0]) or {}
    entry={'name':make_display(r[0],r[1]),
           'days':int(r[4]) if isinstance(r[4],(int,float)) and r[4] else 0,
           'territory':ter,'proptype':attrs.get('proptype',''),
           'month':str(r[16]).strip() if r[16] else '',
           'year':str(r[17]).strip() if r[17] else '','status':'Vacant-Unrented'}
    if r[18]==latest_vac_we: vac_top.append(entry)
    snap={k:v for k,v in entry.items() if k not in ('month','year','status')}
    vac_we_snaps[we_str].append(snap)
vac_top.sort(key=lambda x:-x['days'])
for k in vac_we_snaps: vac_we_snaps[k].sort(key=lambda x:-x['days'])
D2['vac_top']=vac_top; D2['vac_we']=dict(vac_we_snaps)

# Applications
app_rows=read_sheet(FILES['applications'])
app_tbl=defaultdict(lambda:[0,0,0,0])
STATUS_IDX={'Canceled':0,'Denied':1,'Approved':2,'Converted':3}
for r in app_rows[1:]:
    if not r: continue
    addr=r[27] or r[0]; attrs=resolve(str(addr)) if addr else {}
    if not attrs: attrs={}
    status=str(r[8]).strip().title() if r[8] else ''
    ter=attrs.get('territory',''); prop=attrs.get('proptype','')
    year=str(r[30]).strip() if r[30] else ''
    month=str(r[29]).strip() if r[29] else ''
    idx=STATUS_IDX.get(status,-1)
    if idx<0: continue
    for k in agg_keys(year,month,ter,prop): app_tbl[k][idx]+=1
D2['app_donut']={k:list(v) for k,v in app_tbl.items()}

# Listed-to-Leased avg (col AK = pre-computed days)
ltl_tbl=defaultdict(lambda:{'sum':0,'cnt':0})
for r in app_rows[1:]:
    if not r: continue
    days=r[36] if isinstance(r[36],(int,float)) and r[36]>=0 else None
    if days is None: continue
    year=r[37].strftime('%Y') if r[37] and hasattr(r[37],'strftime') else ''
    month=r[37].strftime('%B') if r[37] and hasattr(r[37],'strftime') else ''
    addr=r[27] or r[0]; attrs=resolve(str(addr)) if addr else {}
    ter=attrs.get('territory',''); prop=attrs.get('proptype','')
    for k in agg_keys(year,month,ter,prop):
        ltl_tbl[k]['sum']+=days; ltl_tbl[k]['cnt']+=1
DATA['ltl']={k:round(v['sum']/v['cnt'],1) for k,v in ltl_tbl.items() if v['cnt']>0}

# ── LTL BY AGENT with time keys ───────────────────────────────────────────────
ltl_agent_tbl=defaultdict(lambda:defaultdict(lambda:{'sum':0,'cnt':0}))
for r in app_rows[1:]:
    if not r or r[36] is None: continue
    days=r[36] if isinstance(r[36],(int,float)) and r[36]>=0 else None
    if days is None: continue
    agent=str(r[13]).strip() if r[13] else ''
    if not agent or agent in ('--','Brenden'): continue
    yr=str(r[30]).strip() if r[30] else ''
    mo=str(r[29]).strip() if r[29] else ''
    we=r[31].strftime('%Y-%m-%d') if r[31] and hasattr(r[31],'strftime') else ''
    for k in [f"|||",f"{yr}|||",f"|{mo}||",f"{yr}|{mo}||"]:
        ltl_agent_tbl[agent][k]['sum']+=days; ltl_agent_tbl[agent][k]['cnt']+=1
    if we:
        ltl_agent_tbl[agent][f"||{we}|"]['sum']+=days
        ltl_agent_tbl[agent][f"||{we}|"]['cnt']+=1
DATA['ltl_agent']={a:{k:round(v['sum']/v['cnt'],1) for k,v in kv.items() if v['cnt']>0}
                    for a,kv in ltl_agent_tbl.items()}

# ── PHONE ANSWER % BY AGENT with time keys ────────────────────────────────────
ANSWERED_PH={'Answered','Connected','Accepted','Call connected'}
phone_rows=read_sheet(FILES['phone'])
phone_agent_tbl=defaultdict(lambda:defaultdict(lambda:{'total':0,'ans':0}))
for r in phone_rows[1:]:
    if not r or str(r[1]).strip()!='Incoming': continue
    name=str(r[16]).strip() if r[16] else ''
    role=str(r[17]).strip() if r[17] else ''
    result=str(r[10]).strip() if r[10] else ''
    if not name or role!='Leasing': continue
    yr=str(r[23]).strip() if r[23] else ''
    mo=str(r[22]).strip() if r[22] else ''
    we=r[20].strftime('%Y-%m-%d') if r[20] and hasattr(r[20],'strftime') else ''
    is_ans=result in ANSWERED_PH
    for k in [f"|||",f"{yr}|||",f"|{mo}||",f"{yr}|{mo}||"]:
        phone_agent_tbl[name][k]['total']+=1
        if is_ans: phone_agent_tbl[name][k]['ans']+=1
    if we:
        phone_agent_tbl[name][f"||{we}|"]['total']+=1
        if is_ans: phone_agent_tbl[name][f"||{we}|"]['ans']+=1
DATA['phone_agent']={name:{k:round(v['ans']/v['total']*100,1) if v['total'] else None
                            for k,v in kv.items()}
                      for name,kv in phone_agent_tbl.items()}

# WO status + issues
wo_st=defaultdict(lambda:defaultdict(int))
wo_is=defaultdict(lambda:defaultdict(int))
for r in wo_rows[1:]:
    if not r or not r[0]: continue
    attrs=resolve(r[0]) or {}
    ter=attrs.get('territory',''); prop=attrs.get('proptype','')
    year=str(r[42]).strip() if r[42] else ''
    month=str(r[43]).strip() if r[43] else ''
    status=str(r[10]).strip() if r[10] else ''
    issue=str(r[29]).strip() if r[29] else ''
    for k in agg_keys(year,month,ter,prop):
        if status: wo_st[k][status]+=1
        if issue:  wo_is[k][issue]+=1
D2['wo_status']={k:dict(v) for k,v in wo_st.items()}
D2['wo_issues']={k:dict(sorted(v.items(),key=lambda x:-x[1])[:10]) for k,v in wo_is.items()}

# Concessions
con_rows=read_sheet(FILES['concessions'])
D2['concessions']=[]

# ── CONCESSIONS AGGREGATION (for CEO scorecard filters) ──────────────────────
con_agg=defaultdict(float)
for r in con_rows[1:]:
    if not r or not r[1]: continue
    try: amt=float(str(r[1]).replace(',','').replace('$',''))
    except: amt=0
    if amt==0: continue
    month=str(r[3]).strip() if r[3] else ''
    year =str(r[5]).strip() if r[5] else ''
    prop_addr=str(r[7]).strip() if r[7] else str(r[0]).strip()
    attrs=resolve(prop_addr) or {}
    ter=attrs.get('territory',''); pt=attrs.get('proptype','')
    if ter in SKIP_TERS: ter=''
    for k in agg_keys(year,month,ter,pt): con_agg[k]+=amt
DATA['concessions_agg']={k:round(v,2) for k,v in con_agg.items()}
for r in con_rows[1:]:
    if not r or not r[0]: continue
    attrs=resolve(r[7] or r[0]) or {}
    D2['concessions'].append({'property':str(r[0]).strip(),
        'amount':r[1] if isinstance(r[1],(int,float)) else 0,
        'month':str(r[3]).strip() if r[3] else '',
        'year':str(r[5]).strip() if r[5] else '',
        'we':r[4].strftime('%Y-%m-%d') if r[4] and hasattr(r[4],'strftime') else '',
        'reason':str(r[6]).strip() if r[6] else '',
        'territory':attrs.get('territory',''),'proptype':attrs.get('proptype','')})

# Unpaid (EOM)
up_rows=read_sheet(FILES['unpaid'])
MNAMES2={1:'January',2:'February',3:'March',4:'April',5:'May',6:'June',
         7:'July',8:'August',9:'September',10:'October',11:'November',12:'December'}
up_eom=eom_we_map(up_rows,15,14,16)
up_flat=[]; up_totals=defaultdict(lambda:{'total':0.0,'count':0})
for r in up_rows[1:]:
    if not r or not r[16] or not hasattr(r[16],'strftime'): continue
    eom_key=get_eom_key(r[16],up_eom)
    if not eom_key: continue
    year,month=eom_key
    amt=r[8] if isinstance(r[8],(int,float)) and r[8] and r[8]>0 else 0
    if amt==0: continue
    attrs=resolve(r[1] or r[0]) or {}
    ter=str(r[17]).strip() if r[17] else attrs.get('territory','')
    prop=attrs.get('proptype','')
    up_flat.append({'n':str(r[4]).strip() if r[4] else '','pr':str(r[0]).strip() if r[0] else '',
        'u':str(r[3]).strip() if r[3] else '','a':round(amt,2),
        's':str(r[5]).strip() if r[5] else '','t':ter,'pt':prop,
        'mo':month,'yr':year,'we':r[16].strftime('%Y-%m-%d')})
    for k in agg_keys(year,month,ter,prop):
        up_totals[k]['total']+=amt; up_totals[k]['count']+=1
up_flat.sort(key=lambda x:-x['a'])
D2['unpaid']=up_flat
D2['up_totals']={k:{'total':round(v['total'],2),'count':v['count']} for k,v in up_totals.items()}

# Rent roll
rr_rows=read_sheet(FILES['rent_roll'])
MNAMES3={1:'January',2:'February',3:'March',4:'April',5:'May',6:'June',
         7:'July',8:'August',9:'September',10:'October',11:'November',12:'December'}
rr_eom=eom_we_map(rr_rows,20,19,21)
rr_rent=defaultdict(float); rr_count=defaultdict(int)
for r in rr_rows[1:]:
    if not r or not r[0] or not r[21] or not hasattr(r[21],'strftime'): continue
    eom_key=get_eom_key(r[21],rr_eom)
    if not eom_key: continue
    year,month=eom_key
    attrs=resolve(str(r[0]).strip()) or {}
    ter=attrs.get('territory',''); pt=attrs.get('proptype','')
    if ter in SKIP_TERS: ter=''
    rent=r[9] if isinstance(r[9],(int,float)) and r[9] and r[9]>0 else 0
    for k in agg_keys(year,month,ter,pt):
        rr_rent[k]+=rent; rr_count[k]+=1
D2['rr_rent']={k:round(v,2) for k,v in rr_rent.items()}
D2['rr_count']=dict(rr_count)
D2['rr_avg']={k:round(rr_rent[k]/rr_count[k],2) for k in rr_count if rr_count[k]>0}

# AR%
ar_pct={}
for k in up_totals:
    ar=up_totals[k]['total']; rnt=rr_rent.get(k,0)
    if rnt>0: ar_pct[k]=round(ar/rnt*100,2)
D2['ar_pct']=ar_pct

# 5-star reviews
star_rows=read_sheet(FILES['five_star'],'rawdata reviews')
staff_ter=defaultdict(set)
name_norm={'Anthony Ruiz':'Anthony Andrew Ruiz'}
for r in port_rows[1:]:
    if not r or not r[6]: continue
    ter=str(r[6]).strip()
    for col in [7,8,9,10]:
        if r[col] and str(r[col]).strip() not in ('TBD','TBD_Leasing Agent','Commercial','STONE PM',''):
            staff_ter[str(r[col]).strip()].add(ter)

def norm_plat(p):
    p=str(p).strip().lower()
    if 'yelp' in p and 'oc' in p: return 'Yelp OC'
    if 'google' in p and 'oc' in p: return 'Google OC'
    if 'google' in p and 'sd' in p: return 'Google SD'
    if 'yelp' in p and 'sd' in p: return 'Yelp SD'
    return 'Other'

rev_agg=defaultdict(lambda:defaultdict(lambda:{'c':0,'r':0,'yo':0,'go':0,'gs':0,'ys':0}))
plat_agg=defaultdict(lambda:defaultdict(list))
for r in star_rows[1:]:
    if not r or not r[0]: continue
    name=str(r[0]).strip()
    if name=='TBD': continue
    mapped=name_norm.get(name,name)
    ters=sorted(staff_ter.get(mapped,{''}))
    plat=norm_plat(r[4]) if r[4] else 'Other'
    rating=r[5] if isinstance(r[5],(int,float)) else 5
    year=str(r[3]).strip() if r[3] else ''
    month=str(r[2]).strip() if r[2] else ''
    pk={'Yelp OC':'yo','Google OC':'go','Google SD':'gs','Yelp SD':'ys'}.get(plat,'o')
    ter=ters[0] if ters and ters[0] else ''
    for k in agg_keys(year,month,ter,''):
        rev_agg[k][name]['c']+=1; rev_agg[k][name]['r']+=rating; rev_agg[k][name][pk]+=1
        plat_agg[k][plat].append(rating)
D2['review_lb']={k:sorted([{'n':n,'c':s['c'],'r':round(s['r']/s['c'],2) if s['c'] else 5,
    'yo':s['yo'],'go':s['go'],'gs':s['gs'],'ys':s['ys']} for n,s in names.items()],
    key=lambda x:-x['c'])[:15] for k,names in rev_agg.items()}
D2['plat_ratings']={k:{p:round(sum(v)/len(v),2) for p,v in pd.items()} for k,pd in plat_agg.items()}

# BDM
bdm_rows=read_sheet(FILES['bdm'])
bdm_st=defaultdict(lambda:defaultdict(int))
bdm_src=defaultdict(lambda:defaultdict(int))
bdm_all=[]
for r in bdm_rows[1:]:
    if not r or not r[0]: continue
    status=str(r[3]).strip().lower() if r[3] else ''
    source=str(r[4]).strip() if r[4] else ''
    dv=next((r[i] for i in range(len(r)) if r[i] and hasattr(r[i],'strftime')),None)
    month=dv.strftime('%B') if dv else ''; year=dv.strftime('%Y') if dv else ''
    for k in agg_keys(year,month,'',''):
        bdm_st[k][status]+=1; bdm_src[k][source]+=1
    bdm_all.append({'status':status,'source':source,'month':month,'year':year,'property':str(r[0]).strip()})
D2['bdm_status']={k:dict(v) for k,v in bdm_st.items()}
D2['bdm_source']={k:dict(v) for k,v in bdm_src.items()}
D2['bdm_all']=bdm_all
print(f"  D2: {list(D2.keys())}")


# ── PORT_PT / VAC_PT (proptype breakdown for vacancy filter) ──────────────────
print("[4b/7] Computing port_pt and vac_pt...")

vac_rate_data=read_sheet(FILES['vacancy'],'Vacancy Rate')
ter_map1_data=read_sheet(FILES['vacancy'],'Territory Mapping1')
addr_ter_map={str(r[0]).strip():str(r[1]).strip() for r in ter_map1_data[1:] if r and r[0] and r[1]}

def get_vac_ter_full(r):
    addr=str(r[15]).strip() if r[15] else str(r[0]).strip()
    res=(resolve(addr) or {}).get('territory','')
    if res and res not in SKIP_TERS: return res
    ter=str(r[25]).strip() if r[25] else ''
    if ter and ter not in SKIP_TERS: return ter
    m=addr_ter_map.get(addr,'')
    return m if m not in SKIP_TERS else ''

port_we_ter_pt=defaultdict(lambda:defaultdict(lambda:defaultdict(int)))
for r in port_rows[1:]:
    if not r or not r[13] or not hasattr(r[13],'strftime'): continue
    ter=str(r[6]).strip() if r[6] else ''
    pt=str(r[5]).strip() if r[5] else ''
    if ter in SKIP_TERS: continue
    port_we_ter_pt[r[13].strftime('%Y-%m-%d')][ter][pt]+=1

vac_we_ter_pt=defaultdict(lambda:defaultdict(lambda:defaultdict(int)))
for r in vac_rate_data[1:]:
    if not r or not r[18] or not hasattr(r[18],'strftime'): continue
    if str(r[2]).strip()!='Vacant-Unrented': continue
    ter=get_vac_ter_full(r)
    if ter in SKIP_TERS: continue
    addr=str(r[15]).strip() if r[15] else str(r[0]).strip()
    pt=(resolve(addr) or {}).get('proptype','')
    vac_we_ter_pt[r[18].strftime('%Y-%m-%d')][ter][pt]+=1

D2['port_pt']={k:{t:dict(v) for t,v in tv.items()} for k,tv in port_we_ter_pt.items()}
D2['vac_pt'] ={k:{t:dict(v) for t,v in tv.items()} for k,tv in vac_we_ter_pt.items()}

# ── WEEKLY OVERVIEW SCORECARD ─────────────────────────────────────────────────
print("[4c/7] Computing Weekly Overview scorecard...")

# NPS cumulative up to each WE
nps_rows2=read_sheet(FILES['new_nps'])
nps_raw_we=defaultdict(lambda:{'p':0,'d':0,'t':0})
for r in nps_rows2[1:]:
    if not r or not r[35] or not hasattr(r[35],'strftime'): continue
    p=1 if r[39]==1 else 0; d=1 if r[41]==1 else 0; pa=1 if r[40]==1 else 0
    t=p+pa+d
    if t==0: continue
    nps_raw_we[r[35].strftime('%Y-%m-%d')]['p']+=p
    nps_raw_we[r[35].strftime('%Y-%m-%d')]['d']+=d
    nps_raw_we[r[35].strftime('%Y-%m-%d')]['t']+=t
nps_cum={}; cp=0; cd=0; ct=0; last_nps=None
for we in sorted(nps_raw_we.keys()):
    v=nps_raw_we[we]; cp+=v['p']; cd+=v['d']; ct+=v['t']
    last_nps=round((cp-cd)/ct*100,1) if ct else None
    nps_cum[we]=last_nps

# Unit turn cumulative (blank comment only, include 0-day)
ut_rows2=read_sheet(FILES['unit_turn'],'Unit Turn Details')
ut_raw_we=defaultdict(list)
for r in ut_rows2[1:]:
    if not r or r[17] is None: continue
    if r[16] is not None and str(r[16]).strip()!='': continue
    days=r[17] if isinstance(r[17],(int,float)) else 0
    if r[21] and hasattr(r[21],'strftime'):
        ut_raw_we[r[21].strftime('%Y-%m-%d')].append(days)
ut_cum={}; ut_s=0; ut_c=0
for we in sorted(ut_raw_we.keys()):
    for d in ut_raw_we[we]: ut_s+=d; ut_c+=1
    ut_cum[we]=round(ut_s/ut_c,2) if ut_c else None

# LR% per WE (all WEs, no EOM restriction)
lr_rows2=read_sheet(FILES['lr'])
lr_we=defaultdict(lambda:{'R':0,'M':0})
for r in lr_rows2[1:]:
    if not r or not r[16] or not hasattr(r[16],'strftime'): continue
    status=str(r[10]).strip() if r[10] else ''
    we_str=r[16].strftime('%Y-%m-%d')
    if status=='Renewed': lr_we[we_str]['R']+=1
    elif status in ('Move-Out','Changed to MTM'): lr_we[we_str]['M']+=1
lr_pct_we={we:round(v['R']/(v['R']+v['M'])*100,2) for we,v in lr_we.items() if (v['R']+v['M'])>0}

# AR% per WE = total unpaid / total rent roll at each weekly snapshot
# This matches Power BI which shows the actual weekly AR% (not just EOM)
up_rows2=read_sheet(FILES['unpaid'])
up_by_we2=defaultdict(float)
for r in up_rows2[1:]:
    if not r or not r[16] or not hasattr(r[16],'strftime'): continue
    amt=r[8] if isinstance(r[8],(int,float)) and r[8] and r[8]>0 else 0
    if amt>0: up_by_we2[r[16].strftime('%Y-%m-%d')]+=amt
rr_rows2=read_sheet(FILES['rent_roll'])
rr_by_we2=defaultdict(float)
for r in rr_rows2[1:]:
    if not r or not r[0]: continue
    we=r[21] if len(r)>21 and r[21] and hasattr(r[21],'strftime') else None
    if not we: continue
    rent=r[9] if isinstance(r[9],(int,float)) and r[9] and r[9]>0 else 0
    rr_by_we2[we.strftime('%Y-%m-%d')]+=rent
ar_pct_we2={we:round(up_by_we2[we]/rr_by_we2[we]*100,2)
             for we in up_by_we2 if rr_by_we2.get(we,0)>0}

# Showing conversion from guest card
gc_rows=read_sheet(FILES['guest_card'])
gc_we=defaultdict(lambda:{'total':0,'with_showing':0})
gc_ter_we=defaultdict(lambda:defaultdict(lambda:{'total':0,'with_showing':0}))
for r in gc_rows[1:]:
    if not r or not r[15] or not hasattr(r[15],'strftime'): continue
    we_str=r[15].strftime('%Y-%m-%d')
    gc_we[we_str]['total']+=1
    showings=r[5] if isinstance(r[5],(int,float)) else 0
    if showings>0: gc_we[we_str]['with_showing']+=1
    # Resolve territory from property address (col 11)
    addr=str(r[11]).strip() if r[11] else str(r[0]).strip()
    ter_gc=(resolve(addr) or {}).get('territory','')
    if ter_gc and ter_gc not in SKIP_TERS:
        gc_ter_we[ter_gc][we_str]['total']+=1
        if showings>0: gc_ter_we[ter_gc][we_str]['with_showing']+=1
show_conv_we={we:round(v['with_showing']/v['total']*100,2) for we,v in gc_we.items() if v['total']>0}
show_conv_ter_we={ter:{we:round(v['with_showing']/v['total']*100,2) for we,v in wes.items() if v['total']>0}
                   for ter,wes in gc_ter_we.items()}

# Vacancy $ Loss Rate (Posted=Yes, col AA posted_at, monthly reset, mgmt_fee)
mgmt_fee_map={}
for r in port_rows[1:]:
    if not r or not r[0]: continue
    addr=str(r[0]).strip(); fee=r[14]
    if isinstance(fee,(int,float)) and fee>0:
        mgmt_fee_map[addr]=fee
        mgmt_fee_map[addr.split(',')[0].strip().lower()]=fee
def get_fee(addr):
    addr=str(addr).strip()
    return mgmt_fee_map.get(addr, mgmt_fee_map.get(addr.split(',')[0].strip().lower(), 0.089))

vac_rate3=read_sheet(FILES['vacancy'],'Vacancy Rate')
port_cnt_we=defaultdict(int)
for r in port_rows[1:]:
    if not r or not r[13] or not hasattr(r[13],'strftime'): continue
    ter=str(r[6]).strip() if r[6] else ''
    if ter in SKIP_TERS: continue
    port_cnt_we[r[13].strftime('%Y-%m-%d')]+=1

vac_cnt_we=defaultdict(int)
vac_loss_we=defaultdict(float)
vac_loss_ter_we=defaultdict(lambda:defaultdict(float))
for r in vac_rate3[1:]:
    if not r or not r[18] or not hasattr(r[18],'strftime'): continue
    if str(r[2]).strip()!='Vacant-Unrented': continue
    ter=get_vac_ter_full(r)
    if ter in SKIP_TERS: continue
    we_dt=r[18]; we_str=we_dt.strftime('%Y-%m-%d')
    vac_cnt_we[we_str]+=1
    if str(r[14]).strip()!='Yes': continue
    if not r[26] or not hasattr(r[26],'strftime'): continue
    sched=r[6] if isinstance(r[6],(int,float)) and r[6]>0 else 0
    if sched==0: continue
    prop_addr=str(r[15]).strip() if r[15] else str(r[0]).strip()
    fee=get_fee(prop_addr)
    month_start=we_dt.replace(day=1)
    start=max(r[26], month_start)
    days=(we_dt-start).days+1
    if days<=0: continue
    loss=round(days*(sched/30)*fee,2)
    vac_loss_we[we_str]+=loss
    vac_loss_ter_we[ter][we_str]+=loss

# Vacancy rate per WE
vac_rate_we={we:round(vac_cnt_we[we]/port_cnt_we[we]*100,2) if port_cnt_we.get(we) else 0 for we in vac_cnt_we}

# Build Weekly Overview (scorecard) rows - all territories combined
from datetime import timedelta
all_sc_we=sorted(set(list(vac_cnt_we.keys())+list(nps_cum.keys())+list(lr_pct_we.keys())))
scorecard_rows=[]
last_nps_v=None
for we in all_sc_we:
    we_dt=datetime.datetime.strptime(we,'%Y-%m-%d')
    unit_count=port_cnt_we.get(we,0)
    vacant=vac_cnt_we.get(we,0)
    vac_pct=round(vacant/unit_count*100,2) if unit_count else 0
    if we in nps_cum: last_nps_v=nps_cum[we]
    # Get month name from WE date for rr_avg/concessions lookup
    _SC_MNAMES={1:'January',2:'February',3:'March',4:'April',5:'May',6:'June',
                7:'July',8:'August',9:'September',10:'October',11:'November',12:'December'}
    _we_dt=datetime.datetime.strptime(we,'%Y-%m-%d')
    _yr=str(_we_dt.year); _mo=_SC_MNAMES[_we_dt.month]
    _ru=D2.get('rr_avg',{}).get(f'{_yr}|{_mo}||')  # avg rent/unit for this month
    # WE-level concessions: sum of concession amounts for this specific WE
    _conc=sum(c.get('amount',0) for c in D2.get('concessions',[]) if c.get('we')==we)
    scorecard_rows.append({
        'we':we,'territory':'','proptype':'',
        'unit_count':unit_count,'vacant':vacant,'vac_pct':vac_pct,
        'nps':last_nps_v,
        'opened_wo':0,'rent_unit':_ru,'lr_pct':lr_pct_we.get(we),
        'unit_turn':ut_cum.get(we),'ar_pct':ar_pct_we2.get(we),'concessions':_conc,
    })
D2['scorecard']=scorecard_rows

# Build per-territory scorecard rows (for EOS territory filter)
TERS_SC=['North OC','South OC','SD Properties','Commercial','Brenden','Elderkin','STONE']
for ter in TERS_SC:
    last_nps_ter=None; last_ar_ter=None
    for we in all_sc_we:
        # Vacancy/portfolio from vac_donut2/port_counts we: keys
        we_key='we:'+we
        unit_count=(D2['port_counts'].get(we_key) or {}).get(ter,0)
        vacant    =(D2['vac_donut2'].get(we_key)  or {}).get(ter,0)
        vac_pct   =round(vacant/unit_count*100,2) if unit_count else 0
        # NPS: latest cumulative for this territory up to this WE
        ter_nps = DATA.get('nps_cum_ter',{}).get(ter,{})
        nps_wes_before = [w for w in ter_nps if w<=we]
        if nps_wes_before: last_nps_ter=ter_nps[max(nps_wes_before)]
        # LR%: use territory-level key from DATA.lr_r/lr_m
        lr_key_ter = '|||'  # fallback; territory-level LR not in current data
        # AR%: use territory+month EOM
        we_dt=datetime.datetime.strptime(we,'%Y-%m-%d')
        mo_name={1:'January',2:'February',3:'March',4:'April',5:'May',6:'June',
                 7:'July',8:'August',9:'September',10:'October',11:'November',12:'December'}[we_dt.month]
        yr_str=str(we_dt.year)
        ar_ter=DATA.get('ar_pct',{}).get(f'{yr_str}|{mo_name}|{ter}|')
        if ar_ter is not None: last_ar_ter=ar_ter
        elif last_ar_ter is not None: ar_ter=last_ar_ter
        scorecard_rows.append({
            'we':we,'territory':ter,'proptype':'',
            'unit_count':unit_count,'vacant':vacant,'vac_pct':vac_pct,
            'nps':last_nps_ter,
            'opened_wo':0,'rent_unit':None,
            'lr_pct':lr_pct_we.get(we),
            'unit_turn':ut_cum.get(we),
            'ar_pct':ar_ter if ar_ter else last_ar_ter,
            'concessions':0,
            'show_conv':show_conv_ter_we.get(ter,{}).get(we),
            'vac_loss':vac_loss_ter_we.get(ter,{}).get(we,0) or None,
        })
D2['scorecard']=scorecard_rows  # update with per-territory rows added

# Build EOS rows
from datetime import timedelta as td
eos_rows=[]
last_nps_v=None
for we in all_sc_we:
    we_dt2=datetime.datetime.strptime(we,'%Y-%m-%d')
    prev_dt=we_dt2-td(days=6)
    label=f"{prev_dt.strftime('%-m/%-d')} to {we_dt2.strftime('%-m/%-d')}"
    if we in nps_cum: last_nps_v=nps_cum[we]
    sc=show_conv_we.get(we)
    eos_rows.append({
        'we':we,'label':label,
        'nps':last_nps_v,
        'lr_pct':lr_pct_we.get(we),
        'unit_turn':ut_cum.get(we),
        'ar_pct':ar_pct_we2.get(we),
        'show_conv':sc,
        'vac_loss':round(vac_loss_we.get(we,0),2),
        'vac_rate':vac_rate_we.get(we),
    })
# Carry AR% forward week-over-week (EOM value persists until next EOM)
last_ar_eos=None
for r in eos_rows:
    if r.get('ar_pct') is not None:
        last_ar_eos=r['ar_pct']
    elif last_ar_eos is not None:
        r['ar_pct']=last_ar_eos
D2['eos']=eos_rows
print(f"  scorecard rows: {len(scorecard_rows)}, eos rows: {len(eos_rows)}")


# ── CEO SCORECARD DATA ────────────────────────────────────────────────────────
print("[4d/7] Computing CEO scorecard data...")
try:

  # Lost units by territory from property_directory-Term
  term_rows=read_sheet(FILES['term'])
  lost_ter=defaultdict(int)
  for r in term_rows[1:]:
      if not r or not r[0]: continue
      ter=(resolve(str(r[0]).strip()) or {}).get('territory','')
      if ter and ter not in SKIP_TERS:
          units=int(r[3]) if isinstance(r[3],(int,float)) else 1
          lost_ter[ter]+=units

  # Open WO by territory
  OPEN_ST={'Assigned','Estimate Requested','New','Scheduled','Waiting','Work Done'}
  wo_rows2=read_sheet(FILES['work_orders'])
  wo_open_ter=defaultdict(int)
  for r in wo_rows2[1:]:
      if not r or not r[10]: continue
      if str(r[10]).strip() not in OPEN_ST: continue
      ter=(resolve(r[0]) if r[0] else {}).get('territory','')
      if ter and ter not in SKIP_TERS: wo_open_ter[ter]+=1

  # Phone answer % by territory and by leasing agent
  ANSWERED_SET={'Answered','Connected','Accepted','Call connected'}
  phone_rows2=read_sheet(FILES['phone'])
  phone_ter=defaultdict(lambda:{'total':0,'ans':0})
  staff_ter_map=defaultdict(set)
  # Build staff→territory mapping from portfolio
  for r in port_rows[1:]:
      if not r: continue
      ter=str(r[6]).strip() if r[6] else ''
      if ter in SKIP_TERS: continue
      for col in [7,8,9,10]:
          if r[col] and str(r[col]).strip() not in ('TBD','','Commercial','STONE PM'):
              staff_ter_map[str(r[col]).strip()].add(ter)

  for r in phone_rows2[1:]:
      if not r or str(r[1]).strip()!='Incoming': continue
      name=str(r[16]).strip() if r[16] else ''
      result=str(r[10]).strip() if r[10] else ''
      is_ans=result in ANSWERED_SET
      for ter in staff_ter_map.get(name,set()):
          phone_ter[ter]['total']+=1
          if is_ans: phone_ter[ter]['ans']+=1

  phone_pct_ter={ter:round(v['ans']/v['total']*100,1) if v['total'] else None
                  for ter,v in phone_ter.items()}

  TERS_CEO=['North OC','South OC','SD Properties','Commercial','Brenden','Elderkin','STONE']
  D2['ceo']={
      'latest_we':  D2.get('latest_we',''),
      'ters':       TERS_CEO,
      'phone_ter':  {t:phone_pct_ter.get(t) for t in TERS_CEO},
      'wo_open':    {t:wo_open_ter.get(t,0) for t in TERS_CEO},
      'lost_units': {t:lost_ter.get(t,0) for t in TERS_CEO},
      'concessions':{t:round(DATA.get('concessions_agg',{}).get(f'||{t}|',0),0) for t in TERS_CEO},
      'ltl_agent':  DATA.get('ltl_agent',{}),
      'phone_agent':DATA.get('phone_agent',{}),
  }

  print(f"  CEO data built: phone_ter={phone_pct_ter}")
except Exception as e:
  print(f"  WARNING [4d/7]: {e}")
  D2.setdefault('ceo',{})


# ── CALLS LEADERBOARD & TREND ─────────────────────────────────────────────────
print("[4e/7] Computing calls, churn, WO completion, reviews...")
# Wrap in try/except so any error here doesn't break the full build
try:
  ANSWERED_CALLS={'Answered','Connected','Accepted','Call connected'}
  phone_rows3=read_sheet(FILES['phone'])
  calls_by_person=defaultdict(lambda:{'total':0,'ans':0,'role':''})
  calls_by_month3=defaultdict(lambda:defaultdict(lambda:{'total':0,'ans':0}))
  for r in phone_rows3[1:]:
      if not r or str(r[1]).strip()!='Incoming': continue
      name=str(r[16]).strip() if r[16] else ''
      role=str(r[17]).strip() if r[17] else ''
      result=str(r[10]).strip() if r[10] else ''
      yr=str(r[23]).strip() if r[23] else ''
      mo=str(r[22]).strip() if r[22] else ''
      if not name: continue
      is_ans=result in ANSWERED_CALLS
      calls_by_person[name]['total']+=1
      calls_by_person[name]['role']=role
      if is_ans: calls_by_person[name]['ans']+=1
      if yr and mo:
          calls_by_month3[f"{yr}|{mo}"][name]['total']+=1
          if is_ans: calls_by_month3[f"{yr}|{mo}"][name]['ans']+=1

  calls_lb=[]
  for name,v in calls_by_person.items():
      if v['total']>=50:
          calls_lb.append({'name':name,'role':v['role'],'total':v['total'],
              'ans':v['ans'],'pct':round(v['ans']/v['total']*100,1)})
  calls_lb.sort(key=lambda x:-x['pct'])
  D2['calls_lb']=calls_lb

  calls_trend={}
  ROLES_TRACK={'Property Management','Leasing'}
  for mo_key,people in sorted(calls_by_month3.items()):
      pm_tot=sum(v['total'] for n,v in people.items() if calls_by_person[n]['role']=='Property Management')
      pm_ans=sum(v['ans']   for n,v in people.items() if calls_by_person[n]['role']=='Property Management')
      ls_tot=sum(v['total'] for n,v in people.items() if calls_by_person[n]['role']=='Leasing')
      ls_ans=sum(v['ans']   for n,v in people.items() if calls_by_person[n]['role']=='Leasing')
      calls_trend[mo_key]={'pm_total':pm_tot,'pm_ans':pm_ans,'ls_total':ls_tot,'ls_ans':ls_ans}
  D2['calls_trend']=calls_trend

  # ── WO COMPLETION TIME ────────────────────────────────────────────────────────
  wo_rows3=read_sheet(FILES['work_orders'])
  wo_comp_ter=defaultdict(lambda:{'sum':0,'cnt':0})
  wo_comp_mo=defaultdict(lambda:{'sum':0,'cnt':0})
  for r in wo_rows3[1:]:
      if not r or not r[14] or not r[24]: continue
      if not hasattr(r[14],'strftime') or not hasattr(r[24],'strftime'): continue
      days=(r[24]-r[14]).days
      if days<0 or days>365: continue
      ter=str(r[40]).strip() if r[40] else (resolve(str(r[0]).strip()) or {}).get('territory','')
      if ter and ter not in SKIP_TERS:
          wo_comp_ter[ter]['sum']+=days; wo_comp_ter[ter]['cnt']+=1
      yr=str(r[42]).strip() if r[42] else ''
      mo=str(r[43]).strip() if r[43] else ''
      if yr and mo:
          wo_comp_mo[f"{yr}|{mo}"]['sum']+=days; wo_comp_mo[f"{yr}|{mo}"]['cnt']+=1
  D2['wo_completion']={
      'by_ter':{t:round(v['sum']/v['cnt'],1) for t,v in wo_comp_ter.items() if v['cnt']>0},
      'by_month':{k:round(v['sum']/v['cnt'],1) for k,v in wo_comp_mo.items() if v['cnt']>0}
  }

  # ── REVIEWS DETAIL ────────────────────────────────────────────────────────────
  fs_rows=read_sheet(FILES['five_star'])
  reviews_detail=[]
  PLATFORMS=['YELP OC','GOOGLE OC','Google SD','YELP SD']
  for r in fs_rows[1:]:
      if not r or not r[0]: continue
      entry={'name':str(r[0]).strip(),'total':r[1] if isinstance(r[1],(int,float)) else 0,
             'points':r[6] if len(r)>6 and isinstance(r[6],(int,float)) else 0}
      for i,p in enumerate(PLATFORMS,2):
          entry[p]=r[i] if i<len(r) and isinstance(r[i],(int,float)) else 0
      reviews_detail.append(entry)
  D2['reviews_detail']=reviews_detail

  # ── CHURN DATA ────────────────────────────────────────────────────────────────
  term_rows2=read_sheet(FILES['term'])
  # Print headers to diagnose column layout on live data
  if len(term_rows2)>0:
      print(f"  Term file headers: {[str(h)[:15] for h in term_rows2[0]]}")
  churn_rows=[]
  for r in term_rows2[1:]:
      if not r or not r[0]: continue
      addr=str(r[0]).strip()
      attrs=resolve(addr) or {}
      # Safely get columns with fallbacks
      def safe_col(row, idx, default=''):
          try: return row[idx] if idx < len(row) else default
          except: return default
      units_val = safe_col(r, 3)
      end_val   = safe_col(r, 5)
      reason_val= safe_col(r, 6)
      yr_val    = safe_col(r, 7)
      mo_val    = safe_col(r, 8)
      churn_rows.append({
          'address':addr,
          'units':int(units_val) if isinstance(units_val,(int,float)) else 1,
          'reason':str(reason_val).strip() if reason_val else 'Unknown',
          'end_date':end_val.strftime('%Y-%m-%d') if end_val and hasattr(end_val,'strftime') else '',
          'year':str(yr_val).strip() if yr_val else '',
          'month':str(mo_val).strip() if mo_val else '',
          'territory':attrs.get('territory',''),
          'pm':attrs.get('pm','')
      })
  D2['churn']=churn_rows
  print(f"  Churn rows loaded: {len(churn_rows)}")

  # ── PORTFOLIO COUNT BY TERRITORY (current snapshot for churn %) ───────────────
  # Use latest WE with full territory data (skip partial snapshots with <3 territories)
  we_port_keys=sorted([k for k in D2['port_counts'] if k.startswith('we:')])
  latest_port_we=None
  for k in reversed(we_port_keys):
      if len(D2['port_counts'].get(k,{})) >= 3:
          latest_port_we=k; break
  if not latest_port_we and we_port_keys:
      latest_port_we=we_port_keys[-1]
  D2['port_by_ter_current']=D2['port_counts'].get(latest_port_we,{}) if latest_port_we else {}
  D2['port_latest_we']=latest_port_we.replace('we:','') if latest_port_we else ''
  print(f"  Portfolio snapshot: {latest_port_we} → {D2['port_by_ter_current']}")


  print(f"  calls_lb:{len(calls_lb)} wo_comp:{len(wo_comp_ter)} churn:{len(churn_rows)} reviews:{len(reviews_detail)}")
except Exception as e:
  print(f"  WARNING [4e/7]: {e} — setting empty defaults")
  D2.setdefault('calls_lb',[])
  D2.setdefault('calls_trend',{})
  D2.setdefault('wo_completion',{'by_ter':{},'by_month':{}})
  D2.setdefault('reviews_detail',[])
  D2.setdefault('churn',[])
  D2.setdefault('port_by_ter_current',{})
  D2.setdefault('port_latest_we','')


# ── SHOWMOJO: Leads → Showings → Apps funnel ─────────────────────────────────
print("[4f/7] Computing ShowMojo leads/showings/apps data...")
try:
    def _parse_offmkt_addr(raw):
        raw = str(raw).strip()
        raw = re.sub(r'\s+-\s+\d+\s+bed.*$',  '', raw, flags=re.IGNORECASE).strip()
        raw = re.sub(r'\s+-\s+Studio.*$',       '', raw, flags=re.IGNORECASE).strip()
        raw = re.sub(r',\s*\d+\s+(full|half).*$','', raw, flags=re.IGNORECASE).strip()
        return raw.rstrip(' -').strip()

    def _parse_active_addr(raw):
        raw = str(raw).strip()
        parts = raw.split(',')
        if len(parts) >= 3:
            street  = re.sub(r'\s+-\s+[^,]+$', '', parts[0].strip()).strip()
            city    = parts[1].strip()
            state   = parts[2].strip()
            zipcode = parts[3].strip() if len(parts) > 3 else ''
            return f"{street} {city}, {state} {zipcode}".strip(), parts[0].strip()
        return raw, raw

    def _sm_resolve(parsed):
        r = resolve(parsed)
        if r: return r
        parts = parsed.split()
        if parts:
            num = parts[0]
            snu = re.sub(r'\s*-\s*[a-z0-9/]+\s*$', '', parsed.lower()).strip()
            for addr, attrs in master_full.items():
                if addr.startswith(num+' ') or addr.startswith(num+'-'):
                    if snu[:20] in addr.lower() or addr.lower()[:20] in snu:
                        return attrs
        return None

    def _sm_resolve_canon(parsed):
        """Returns (attrs, canonical_portfolio_address)"""
        if parsed in master_full: return master_full[parsed], parsed
        short = parsed.split(',')[0].strip().lower()
        if short in master_short: return master_full[master_short[short][0]], master_short[short][0]
        parts = parsed.split()
        if parts:
            num = parts[0]
            snu = re.sub(r'\s*-\s*[a-z0-9/]+\s*$', '', parsed.lower()).strip()
            for addr, attrs in master_full.items():
                if addr.startswith(num+' ') or addr.startswith(num+'-'):
                    if snu[:20] in addr.lower() or addr.lower()[:20] in snu:
                        return attrs, addr
        return None, None

    sm_active_rows = read_sheet(FILES['sm_active'])
    sm_offmkt_rows = read_sheet(FILES['sm_offmkt'])
    app_rows3      = read_sheet(FILES['applications'])

    print(f"  SM files: {len(sm_offmkt_rows)-1} off-mkt, {len(sm_active_rows)-1} active, {len(app_rows3)-1} apps")

    # ── Off-market: single cumulative snapshot (no WE column) ─────────────────
    # col: 0=listing 1=leads 2=sched_showings 3=actual_showings 4=actual_conv% 7=dom
    sm_by_ter       = defaultdict(lambda:{'leads':0,'sched':0,'actual':0,'dom':[],'listings':0})
    sm_addrs_by_ter = defaultdict(set)   # canonical portfolio addr → territory
    SM_MISMATCHES   = []

    for r in sm_offmkt_rows[1:]:
        if not r or not r[0]: continue
        parsed = _parse_offmkt_addr(str(r[0]))
        attrs, canon = _sm_resolve_canon(parsed)
        if not attrs or not attrs.get('territory') or attrs['territory'] in SKIP_TERS:
            SM_MISMATCHES.append({'source':'off_market','raw':str(r[0]),'parsed':parsed})
            continue
        ter    = attrs['territory']
        leads  = int(r[1]) if isinstance(r[1],(int,float)) else 0
        sched  = int(r[2]) if isinstance(r[2],(int,float)) else 0
        actual = int(r[3]) if isinstance(r[3],(int,float)) else 0
        dom    = float(r[7]) if isinstance(r[7],(int,float)) and r[7]>0 else 0
        d = sm_by_ter[ter]
        d['leads']+=leads; d['sched']+=sched; d['actual']+=actual; d['listings']+=1
        if dom>0: d['dom'].append(dom)
        sm_addrs_by_ter[ter].add(canon)

    # ── Active listings: current snapshot ────────────────────────────────────
    sm_active_ter      = defaultdict(lambda:{'leads':0,'sched':0,'actual':0,'dom':[],'listings':0})
    sm_active_listings = []

    for r in sm_active_rows[1:]:
        if not r or not r[0]: continue
        raw    = str(r[0]).strip()
        p1, p2 = _parse_active_addr(raw)
        attrs  = _sm_resolve(p1) or _sm_resolve(p2)
        we_raw = r[14]
        we_str = we_raw.strftime('%Y-%m-%d') if hasattr(we_raw,'strftime') else str(we_raw)[:10]
        if not attrs or not attrs.get('territory') or attrs['territory'] in SKIP_TERS:
            SM_MISMATCHES.append({'source':'active','raw':raw,'parsed':p1,'we':we_str})
            continue
        ter    = attrs['territory']
        leads  = int(r[3]) if isinstance(r[3],(int,float)) else 0
        sched  = int(r[2]) if isinstance(r[2],(int,float)) else 0
        actual = int(r[1]) if isinstance(r[1],(int,float)) else 0
        dom    = int(r[4]) if isinstance(r[4],(int,float)) else 0
        d = sm_active_ter[ter]
        d['leads']+=leads; d['sched']+=sched; d['actual']+=actual; d['listings']+=1
        if dom>0: d['dom'].append(dom)
        disp_addr = re.sub(r',\s*[A-Z]{2},\s*\d{5}.*$','',p1).strip()
        s2s_a = round(actual/sched*100,1) if sched>0 else None
        sm_active_listings.append({'address':disp_addr,'territory':ter,
            'leads':leads,'sched':sched,'actual':actual,'dom':dom,'s2s_pct':s2s_a})
    sm_active_listings.sort(key=lambda x:(x['territory'],-x['leads']))

    # ── Apps received for ShowMojo-tracked properties ─────────────────────────
    # Show→App % = Apps Received (all statuses) / Actual Showings
    # Scoped to properties appearing in the ShowMojo off-market file
    apps_by_ter = defaultdict(int)
    for r in app_rows3[1:]:
        if not r or not r[27]: continue
        attrs_a, canon_a = _sm_resolve_canon(str(r[27]).strip())
        if not attrs_a or not attrs_a.get('territory') or attrs_a['territory'] in SKIP_TERS: continue
        ter_a = attrs_a['territory']
        if canon_a in sm_addrs_by_ter.get(ter_a, set()):
            apps_by_ter[ter_a] += 1

    def _sm_fin(d, ter):
        avg_dom = round(sum(d['dom'])/len(d['dom']),1) if d.get('dom') else None
        l2s     = round(d['actual']/d['leads']*100,1)  if d.get('leads',0)>0  else None
        s2s     = round(d['actual']/d['sched']*100,1)  if d.get('sched',0)>0  else None
        apps_n  = apps_by_ter.get(ter, 0) if ter else 0
        s2a     = round(apps_n/d['actual']*100,1)      if d.get('actual',0)>0 else None
        return {'leads':d['leads'],'sched':d['sched'],'actual':d['actual'],
                'avg_dom':avg_dom,'l2s_pct':l2s,'s2s_pct':s2s,
                'show_to_app_pct':s2a,'apps_received':apps_n,'listings':d['listings']}

    def _sm_fin_active(d):
        avg_dom = round(sum(d['dom'])/len(d['dom']),1) if d.get('dom') else None
        s2s     = round(d['actual']/d['sched']*100,1)  if d.get('sched',0)>0  else None
        return {'leads':d['leads'],'sched':d['sched'],'actual':d['actual'],
                'avg_dom':avg_dom,'s2s_pct':s2s,'listings':d['listings']}

    D2['showmojo'] = {
        'by_ter_all':      {ter:_sm_fin(d,ter) for ter,d in sm_by_ter.items()},
        'active_ter':      {ter:_sm_fin_active(d) for ter,d in sm_active_ter.items()},
        'active_listings':  sm_active_listings,
        'mismatches':       SM_MISMATCHES,
    }

    total_om = sum(d['listings'] for d in sm_by_ter.values())
    total_ac = sum(d['listings'] for d in sm_active_ter.values())
    print(f"  SM matched: {total_om} off-mkt + {total_ac} active listings")
    print(f"  SM apps by ter: { {t:n for t,n in apps_by_ter.items()} }")
    if SM_MISMATCHES:
        print(f"  ⚠️  ShowMojo mismatches ({len(SM_MISMATCHES)}):")
        for m in SM_MISMATCHES:
            print(f"    [{m['source']}] '{m['raw']}'")
    else:
        print("  ✅ All ShowMojo addresses matched to portfolio")
except Exception as e:
    import traceback; traceback.print_exc()
    print(f"  WARNING [4f/7]: ShowMojo error: {e}")
    D2.setdefault('showmojo', {'by_ter_all':{},'active_ter':{},'active_listings':[],'mismatches':[]})



print("[5/7] Reading HTML template...")
if not os.path.exists(TEMPLATE_PATH):
    print(f"ERROR: Template not found at {TEMPLATE_PATH}")
    sys.exit(1)
with open(TEMPLATE_PATH, encoding='utf-8') as f:
    template = f.read()
print(f"  Template: {len(template):,} chars")

print("[6/7] Assembling dashboard...")
dashboard = (template
    .replace('<<<DATA_PLACEHOLDER>>>', 'const DATA = ' + json.dumps(DATA) + ';')
    .replace('<<<D2_PLACEHOLDER>>>',   'const D2 = '   + json.dumps(D2)   + ';')
)

if '<<<DATA_PLACEHOLDER>>>' in dashboard or '<<<D2_PLACEHOLDER>>>' in dashboard:
    print("ERROR: Placeholders not replaced!")
    sys.exit(1)

print("[7/7] Writing output...")
out_path = os.path.join(OUTPUT_DIR, 'index.html')
with open(out_path, 'w', encoding='utf-8') as f:
    f.write(dashboard)

build_date = datetime.datetime.now().strftime('%Y-%m-%d %H:%M UTC')
size_kb = len(dashboard)/1024
print("=" * 60)
print(f"Done!  Built: {build_date}")
print(f"Output: {out_path} ({size_kb:.0f} KB)")
print("=" * 60)
