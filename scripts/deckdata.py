import pickle, json, re, datetime as dt, pandas as pd, numpy as np
M=pickle.load(open('model.pkl','rb')); E=pickle.load(open('evm.pkl','rb'))
df,preds,parent,children=pickle.load(open('sched.pkl','rb'))
L=M['L']; d=df.set_index('ID')
out={}
# logic stats
rel={'FS':0,'SS':0,'FF':0,'SF':0}; lag=0; nlinks=0
succ=set()
for i,lst in preds.items():
    for p,t,l in lst: rel[t]+=1; nlinks+=1; lag+= (l!=0); succ.add(p)
leaf=set(L.ID)
out['links']=nlinks; out['rel']=rel; out['lagged']=lag
out['no_pred']=len([i for i in leaf if not preds[i]]); out['no_succ']=len([i for i in leaf if i not in succ])
out['maxdur']=int(L.dur.max()); out['gt20']=int((L.dur>20).sum())
# phases
ph=[]
for w,name in [('1','Pre-construction & mobilisation'),('2','Substructure (4 zones)'),('3','Superstructure – 12 towers MIVAN'),('4','Finishing works'),('5','MEP & lifts'),('6','External development'),('7','T&C, OC & handover')]:
    r=df[df.WBS==w].iloc[0]; ph.append([name,str(r.Start.date()),str(r.Finish.date())])
out['phases']=ph
def fd(i,col='Finish'): return str(d.loc[i,col].date())
ms=[('M0','Notice to proceed / mobilisation','2024-01-01','Contractual','Client / NBCC'),
    ('M1','MIVAN mock-up approved & formwork ready',str(df[df.WBS=='1.8'].iloc[0].Finish.date()),'Internal','Formwork vendor'),
    ('M2','Bulk excavation complete (all zones)',str(df[df.WBS=='2.1'].iloc[0].Finish.date()),'Internal','Substructure'),
    ('M3','Raft foundation complete',str(df[df.WBS=='2.3'].iloc[0].Finish.date()),'Internal','Substructure'),
    ('M4','Podium / ground slab complete – towers released',str(df[df.WBS=='2.10'].iloc[0].Finish.date()),'Internal','Substructure'),
    ('M5','First tower topped out (T01 Krishna)',str(df[df.WBS=='T01'].iloc[0].Finish.date()),'Internal','Superstructure'),
    ('M6','All 12 towers topped out (T12 Gomati)',str(df[df.WBS=='3'].iloc[0].Finish.date()),'Contractual','Superstructure'),
    ('M7','Lifts commissioned – all towers',str(df[df.WBS=='5.13'].iloc[0].Finish.date()),'Internal','MEP / OEM'),
    ('M8','Finishing complete – all towers',str(df[df.WBS=='4'].iloc[0].Finish.date()),'Internal','Finishing'),
    ('M9','Occupancy certificate – last tower',fd(2751),'Contractual','Liaison'),
    ('M10','Final handover to owners association',fd(2758),'Contractual','Client')]
out['milestones']=ms
# first handover
out['first_handover']=fd(2679)
# floor cycle package T01 F10
pk=[]
for s in 'abcde':
    r=L[L.WBS=='T01.F10.'+s].iloc[0]
    res=', '.join('%s %g'%(n.replace(' (Unskilled)',''),u) for n,u in r.crew if M['RES'][[x[0] for x in M['RES']].index(n)][3]!='Material')
    mats=', '.join('%s %.1f %s'%(n,q,M['MAT'][n][0]) for n,q in r.mats)
    pk.append([r.WBS,r['Task Name'].split('– ',2)[-1][:60],int(r.dur),r.Predecessors,(res+('; '+mats if mats else '')),round(r.total/1e5,1)])
out['pkg']=pk
# quarterly PV
mc=M['m_cost']/1e7
q=mc.groupby(mc.index.asfreq('Q')).sum()
out['qpv']=[[str(k).replace('Q','-Q'),round(v,1)] for k,v in q.items()]
out['monthly']=[[p.strftime('%b-%y'),round(float(v),2)] for p,v in mc.items()]
out['head']=[[p.strftime('%b-%y'),int(round(v))] for p,v in M['m_head'].items()]
out['peak_head_month']=M['m_head'].idxmax().strftime('%b %Y'); out['peak_head']=int(round(M['m_head'].max()))
out['peak_day']=int(M['head'].max())
wd=pd.Series(M['head']).sum(); out['worker_months']=int(wd/26)
out['avg_head']=int(round(M['m_head'][M['m_head']>10].mean()))
# cost by phase & type
PH={'1':'Pre-construction','2':'Substructure','3':'Superstructure','4':'Finishing','5':'MEP & lifts','6':'External dev.','7':'T&C & handover'}
g=L.groupby('phase')[['lab','equip','mat','fixed','total']].sum()/1e7
out['phase_cost']=[[PH[p],round(g.loc[p,'total'],1)] for p in PH]
bac=L.total.sum()/1e7; out['bac']=round(bac,2)
boq=sum(x['amt'] for x in M['items'])/1e7
ext=sum(a for _,a,_ in M['EXT'])/1e7
oh=(boq+ext)*15/115
out['types']={'Material':round(L.mat.sum()/1e7,1),'Specialist subcontract & balance of rates':round(L.fixed.sum()/1e7-oh,1),
              'Contractor OH&P (in DSR rates)':round(oh,1),'Labour & staff':round(L.lab.sum()/1e7,1),'Equipment':round(L.equip.sum()/1e7,1)}
out['boq']=round(boq,2); out['ext']=round(ext,2)
out['allow']=round(sum(q*r for _,_,_,q,r,_ in M['ALLOW'])/1e7,2)
out['psf']=round(bac*1e7/(558449.55*10.7639))
out['monthly_peak']=[mc.idxmax().strftime('%b %Y'), round(float(mc.max()),1)]
y1=mc[mc.index.year==2024].sum(); y2=mc[mc.index.year==2025].sum(); y3=mc[mc.index.year==2026].sum()
out['years']=[round(y1/bac*100),round(y2/bac*100),round(y3/bac*100)]
# resources
R=M['RES']; out['nres']={g_:sum(1 for r in R if r[3]==g_) for g_ in ['Labour','Staff','Equipment','Material']}
out['nassign']=int(sum(len(c) for c in L.crew)+sum(len(m) for m in L.mats))
# key material quantities
mq={}
for _,r in L.iterrows():
    for n,qq in r.mats: mq[n]=mq.get(n,0)+qq
out['mq']={k:round(v) for k,v in mq.items()}
# float
out['tf']={'zero':int((L.TF<=0).sum()),'le10':int((L.TF<=10).sum()),'le30':int((L.TF<=30).sum()),'gt60':int((L.TF>60).sum())}
ext6=L[L.phase=='6'].TF; out['ext_float']=[int(ext6.min()),int(ext6.max())]
# evm
out['evm']={k:(round(float(v),2) if not isinstance(v,dict) and not isinstance(v,list) else None) for k,v in E.items() if k in ('eac','cpi','spi','ev','ac','bac')}
st=pd.Period('2025-04','M'); pv=float((M['m_cost'].loc[:st].sum())/1e7); out['evm']['pv']=round(pv,2)
out['evm_rows']=[[p.strftime('%b-%y'),round(pv_,2),(round(ev,2) if ev is not None else None),(round(ac,2) if ac is not None else None),(round(fc,2) if fc is not None else None)] for p,pv_,ev,ac,fc in E['EVM_ROWS']]
# cumulative PV at status
out['driving']=197
json.dump(out,open('deck.json','w'),indent=1,default=str)
print(json.dumps({k:v for k,v in out.items() if k not in ('monthly','head','evm_rows','qpv')},indent=1,default=str))
print(out['qpv'])
