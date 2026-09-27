import pandas as pd, re, datetime as dt, pickle
SRC='../inputs/Golf_Homes_PhII_L3_Schedule_GRAP.xlsx'
df=pd.read_excel(SRC,'L3_Schedule',dtype={'WBS':str})
df['dur']=df.Duration.str.extract(r'(\d+)').astype(int)
df['leaf']=df['Resource / Crew'].notna()
def lvl(r):
    if r.WBS=='0': return 1
    if r.leaf: return 4
    if re.fullmatch(r'\d+',r.WBS): return 2
    return 3
df['OL']=df.apply(lvl,axis=1)
START=dt.date(2024,1,1)
def grap(d):  # d = date
    return (d.month==11 and d.day>=16) or d.month==12 or (d.month==1 and d.day<=15)
def cat(r): return str(r['GRAP Calendar Category'])[:1]
# day index 0 = 2024-01-01
N=1500
days=[START+dt.timedelta(i) for i in range(N)]
cap={'A':[0 if grap(d) else 1 for d in days],'B':[0.6 if grap(d) else 1 for d in days],'C':[1]*N}
work={k:[x>0 for x in v] for k,v in cap.items()}
import numpy as np
cum={k:np.cumsum(v) for k,v in work.items()}
def next_work(c,i):
    while not work[c][i]: i+=1
    return i
def finish_from_start(c,s,d):  # s is work day idx (start of day), returns last working day idx
    i=s; n=cap[c][s]
    while n<d-1e-9:
        i+=1
        n+=cap[c][i]
    return i
def add_work(c,i,lag): # move lag working days forward from start of day i
    n=0
    while n<lag:
        i+=1
        if work[c][i]: n+=1
    return i
preds={}
for _,r in df.iterrows():
    p=r.Predecessors
    L=[]
    if isinstance(p,str) and p.strip():
        for tok in p.split(','):
            m=re.fullmatch(r'(\d+)(FS|SS|FF|SF)?([+-]\d+)?',tok.strip())
            L.append((int(m.group(1)),m.group(2) or 'FS',int(m.group(3) or 0)))
    preds[r.ID]=L
rows={r.ID:r for _,r in df.iterrows()}
ES={};EF={}
children={}
# parent mapping by outline
stack=[]
parent={}
for _,r in df.iterrows():
    while stack and rows[stack[-1]].OL>=r.OL: stack.pop()
    parent[r.ID]=stack[-1] if stack else None
    stack.append(r.ID)
    if parent[r.ID]: children.setdefault(parent[r.ID],[]).append(r.ID)
def leaves_under(i):
    out=[]
    for c in children.get(i,[]):
        out += [c] if rows[c].leaf else leaves_under(c)
    return out
import sys; sys.setrecursionlimit(10000)
memo={}
def calc(i):
    if i in ES: return
    r=rows[i]
    if not r.leaf:
        ls=leaves_under(i)
        for l in ls: calc(l)
        ES[i]=min(ES[l] for l in ls); EF[i]=max(EF[l] for l in ls); return
    c=cat(r); s=0
    for (p,t,lag) in preds[i]:
        calc(p)
        if t=='FS': cand=EF[p]+1
        elif t=='SS': cand=ES[p]
        else: raise Exception(t)
        cand = cand + lag if lag else cand   # calendar-day lags (elapsed) assumption
        s=max(s,cand)
    s=next_work(c,s)
    ES[i]=s; EF[i]=finish_from_start(c,s,r.dur)
for i in rows: calc(i)
df['cES']=[days[ES[i]] for i in df.ID]; df['cEF']=[days[EF[i]] for i in df.ID]
df['dS']=(pd.to_datetime(df.cES)-df.Start).dt.days
df['dF']=(pd.to_datetime(df.cEF)-df.Finish).dt.days
print(df.dS.describe(), df.dF.describe())
print(df[(df.dS!=0)|(df.dF!=0)][['ID','WBS','Duration','Predecessors','GRAP Calendar Category','Start','cES','Finish','cEF']].head(30).to_string())
print((df.dS!=0).sum(),(df.dF!=0).sum())
pickle.dump((df,preds,parent,children),open('sched.pkl','wb'))

# ---------- backward pass ----------
succ={}
for i,L in preds.items():
    for (p,t,lag) in L: succ.setdefault(p,[]).append((i,t,lag))
leafs=[i for i in rows if rows[i].leaf]
PEND=max(EF[i] for i in leafs)
def prev_work(c,i):
    while not work[c][i]: i-=1
    return i
def start_from_finish(c,f,d):
    i=f; n=cap[c][f]
    while n<d-1e-9:
        i-=1; n+=cap[c][i]
    return i
LS={};LF={}
# summary predecessors? expand: if a pred is a summary, treat links via its leaves (not present check)
assert all(rows[p].leaf for L in preds.values() for (p,_,_) in L)
order=sorted(leafs,key=lambda i:-EF[i])
def lcalc(i):
    if i in LS: return
    c=cat(rows[i]); lf=PEND; ls_lim=None
    for (s,t,lag) in succ.get(i,[]):
        lcalc(s)
        if t=='FS': lf=min(lf, LS[s]-lag-1)
        elif t=='SS':
            v=LS[s]-lag
            ls_lim=v if ls_lim is None else min(ls_lim,v)
    lf=prev_work(c,lf)
    ls=start_from_finish(c,lf,rows[i].dur)
    if ls_lim is not None and ls_lim<ls:
        ls=next_work(c,ls_lim) if work[c][ls_lim] else prev_work(c,ls_lim)
        lf=finish_from_start(c,ls,rows[i].dur)
    LS[i]=ls; LF[i]=lf
for i in order: lcalc(i)
df['TF']=[ (LS[i]-ES[i]) if rows[i].leaf else None for i in df.ID]
df['crit']=df.TF.apply(lambda x: x is not None and x<=0)
print('project end',days[PEND],'crit leaves',df.crit.sum(),'of',df.leaf.sum())
print('near-crit <=10',((df.TF>0)&(df.TF<=10)).sum())
print(df[df.crit].groupby(df.WBS.str.split('.').str[0]).size())
print(df[df.leaf].TF.describe())
pickle.dump((df,preds,parent,children),open('sched.pkl','wb'))
