from __future__ import annotations

import io
import json
import re
import unicodedata
from html import escape
from datetime import datetime
from pathlib import Path

import gspread
import numpy as np
import pandas as pd
import plotly.express as px
import streamlit as st
from google.oauth2.service_account import Credentials
from streamlit_autorefresh import st_autorefresh


APP_DIR = Path(__file__).resolve().parent
DEMO_FILE = APP_DIR / "data" / "inventaires_demo.xlsx"
DEMO_JSON = APP_DIR / "data" / "inventaires_demo.json"
ZONE_SHEETS = ["BINGA", "BONGANDANGA", "BOSOMANZI", "BOSOMONDANDA", "BOSONDJO", "LISALA", "PIMU"]
VACCINES = {
    "BCG": 18, "VPO": 21, "DTC-HépB-Hib": 24, "PCV-13": 27,
    "Rotarix": 30, "VPI": 33, "VAR": 36, "VAA": 39, "Td": 42, "VAP": 45,
}

st.set_page_config(page_title="Chaîne du froid PEV", page_icon="❄️", layout="wide")
st.markdown("""
<style>
:root{--navy:#092c4c;--ice:#e8f7fb;--cyan:#0891b2;--green:#059669;--amber:#d97706;--red:#dc2626;--line:#d7e7ef}
.stApp{background:radial-gradient(circle at 90% 0,rgba(6,182,212,.12),transparent 27rem),linear-gradient(#f8fcfe,#f3f8fb)}
.block-container{max-width:1500px;padding-top:1rem}.hero{position:relative;overflow:hidden;background:linear-gradient(120deg,#092c4c,#086788 58%,#079b88);padding:1.6rem 1.8rem;border-radius:22px;color:white;box-shadow:0 15px 36px rgba(9,44,76,.2)}
.hero:after{content:"❄";position:absolute;right:2rem;top:-1.8rem;font-size:9rem;color:rgba(255,255,255,.09)}.hero small{color:#a5f3fc;font-weight:800;letter-spacing:.16em}.hero h1{color:white;margin:.2rem 0;font-size:2.15rem}.hero p{color:#dff8ff;margin:.35rem 0 0}
.cards{display:grid;grid-template-columns:repeat(6,minmax(0,1fr));gap:11px;margin:1rem 0}.card{background:white;border:1px solid var(--line);border-top:4px solid var(--accent);border-radius:16px;padding:.9rem;box-shadow:0 6px 18px rgba(25,70,100,.07)}.label{font-size:.75rem;color:#5d7185;font-weight:750;min-height:2.15em}.value{font-size:1.55rem;color:#10253d;font-weight:850}.hint{font-size:.68rem;color:var(--accent);font-weight:700}
.module-note{background:linear-gradient(90deg,#e9f8fd,#edfbf7);border:1px solid #cce8ef;border-left:5px solid var(--cyan);border-radius:12px;padding:.7rem .9rem;margin:.5rem 0;color:#274b63}
.decision-grid{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:12px;margin:.8rem 0 1rem}.decision{background:#fff;border:1px solid var(--line);border-radius:16px;padding:1rem;box-shadow:0 6px 18px rgba(25,70,100,.07)}.decision h4{margin:0 0 .35rem;color:#183b56}.decision p{margin:.2rem 0;color:#52677a;font-size:.86rem}.badge{display:inline-block;border-radius:999px;padding:.22rem .55rem;font-size:.72rem;font-weight:800}.critical{background:#fee2e2;color:#991b1b}.high{background:#ffedd5;color:#9a3412}.watch{background:#fef9c3;color:#854d0e}.good{background:#dcfce7;color:#166534}
.insight{border-radius:13px;padding:.8rem 1rem;margin:.45rem 0;border-left:5px solid var(--accent);background:#fff;box-shadow:0 3px 12px rgba(25,70,100,.05)}.insight strong{color:#183b56}.insight span{color:#52677a}.legend-box{background:#f8fafc;border:1px solid var(--line);border-radius:12px;padding:.65rem .8rem;color:#52677a;font-size:.78rem}
[data-testid="stSidebar"]{background:linear-gradient(#fff,#edf8fb);border-right:1px solid var(--line)}div[data-testid="stRadio"] label{background:white;border:1px solid #d7e5ec;border-radius:999px;padding:.3rem .55rem}div[data-testid="stRadio"] label:has(input:checked){background:#dff7fa;border-color:#22b8cf;font-weight:700;color:#075985}
div[data-testid="stPlotlyChart"],div[data-testid="stDataFrame"]{background:white;border:1px solid var(--line);border-radius:16px;padding:.2rem;box-shadow:0 5px 16px rgba(25,70,100,.06)}.stDownloadButton button,.stButton button{border-radius:11px;min-height:2.6rem;font-weight:700}
@media(max-width:950px){.cards{grid-template-columns:repeat(2,1fr)}.decision-grid{grid-template-columns:1fr}.hero h1{font-size:1.55rem}}@media(max-width:520px){.block-container{padding:.55rem}.cards{gap:7px}.card{padding:.7rem}.value{font-size:1.25rem}.hero{padding:1.15rem}.hero h1{font-size:1.35rem}}
</style>
""", unsafe_allow_html=True)


def norm(value: object) -> str:
    text = unicodedata.normalize("NFKD", str(value or "").strip().lower())
    return "".join(c for c in text if not unicodedata.combining(c))


def number(value: object) -> float:
    if value is None or value == "": return np.nan
    if isinstance(value, (int, float)): return float(value)
    text = str(value).replace("\u00a0", "").replace(" ", "").replace(",", ".")
    pct = text.endswith("%")
    try: result = float(text.rstrip("%"))
    except ValueError: return np.nan
    return result / 100 if pct else result


def yes(value: object) -> bool:
    return norm(value) in {"oui", "yes", "o", "1", "vrai", "present", "disponible"}


@st.cache_data(ttl=300, show_spinner="Synchronisation avec Google Sheets…")
def load_google(sheet_id: str, service_info_json: str) -> tuple[dict[str, list[list]], str]:
    info = json.loads(service_info_json)
    credentials = Credentials.from_service_account_info(info, scopes=["https://www.googleapis.com/auth/spreadsheets.readonly", "https://www.googleapis.com/auth/drive.readonly"])
    book = gspread.authorize(credentials).open_by_key(sheet_id)
    matrices = {name: book.worksheet(name).get_all_values(value_render_option="UNFORMATTED_VALUE") for name in ZONE_SHEETS}
    return matrices, datetime.now().astimezone().strftime("%Y-%m-%d %H:%M:%S %Z")


@st.cache_data(show_spinner=False)
def load_demo() -> dict[str, list[list]]:
    return json.loads(DEMO_JSON.read_text(encoding="utf-8"))


def google_config():
    try:
        cfg = st.secrets.get("google_sheet", {})
        account = st.secrets.get("google_service_account", {})
        sheet_id = str(cfg.get("spreadsheet_id", "")).strip()
        if account and account.get("client_email") and sheet_id:
            return sheet_id, json.dumps(dict(account))
    except (FileNotFoundError, KeyError): pass
    return None


def cell(row: list, index: int):
    return row[index] if index < len(row) else None


def parse(matrices: dict[str, list[list]]) -> tuple[pd.DataFrame, pd.DataFrame]:
    sites, stocks = [], []
    for zone, matrix in matrices.items():
        for raw in matrix[3:]:
            area = str(cell(raw, 1) or "").strip()
            pop = number(cell(raw, 2))
            if not area or norm(area) in {"total", "totaux"} or pd.isna(pop): continue
            fridge_count = number(cell(raw, 79)); status = norm(cell(raw, 80))
            broken = "panne" in status and not ("non" in status and "panne" in status)
            functional = any(x in status for x in ["fonctionnel", "bon etat", "marche"]) and not broken
            sites.append({
                "zone":zone.title(), "aire_sante":area, "population":pop,
                "supervisions_prevues":number(cell(raw,3)), "supervisions_realisees":number(cell(raw,4)),
                "seances_planifiees":sum(np.nan_to_num([number(cell(raw,i)) for i in (9,10,11)])),
                "seances_realisees":sum(np.nan_to_num([number(cell(raw,i)) for i in (12,14,16)])),
                "refrigerateurs":fridge_count, "statut_refrigerateur":str(cell(raw,80) or "Non renseigné"),
                "refrigerateur_en_panne":broken, "refrigerateur_fonctionnel":functional,
                "date_debut_panne":cell(raw,81), "fridge_tag":yes(cell(raw,82)), "maintenance_preventive":yes(cell(raw,83)),
                "sites_surveillance":number(cell(raw,74)), "visites_surveillance":sum(np.nan_to_num([number(cell(raw,i)) for i in (75,76,77,78)])),
            })
            for vaccine, start in VACCINES.items():
                need, available, weeks = number(cell(raw,start)), number(cell(raw,start+1)), number(cell(raw,start+2))
                if pd.isna(weeks) and need and not pd.isna(available): weeks = available / need
                stocks.append({"zone":zone.title(),"aire_sante":area,"vaccin":vaccine,"besoin_hebdomadaire":need,"stock_disponible":available,"couverture_semaines":weeks,"rupture":bool(not pd.isna(available) and available <= 0)})
    return pd.DataFrame(sites), pd.DataFrame(stocks)


def csv_bytes(df): return df.to_csv(index=False).encode("utf-8-sig")
def pct(x): return "N/D" if pd.isna(x) else f"{x:.1%}".replace(".",",")
def integer(x): return "N/D" if pd.isna(x) else f"{int(round(x)):,}".replace(","," ")
def style(fig, height=410):
    fig.update_layout(template="plotly_white",height=height,margin=dict(l=15,r=15,t=55,b=25),paper_bgcolor="rgba(0,0,0,0)",font=dict(family="Arial",color="#334155"),legend_title_text="")
    fig.update_xaxes(gridcolor="#eaf1f5"); fig.update_yaxes(gridcolor="#eaf1f5"); return fig


st.markdown('<section class="hero"><small>ANTENNE PEV LISALA</small><h1>Tableau de bord de la chaîne du froid</h1><p>Disponibilité des vaccins, fonctionnalité des équipements, maintenance et continuité des services.</p></section>', unsafe_allow_html=True)
cfg = google_config(); st_autorefresh(interval=5*60*1000,key="refresh_google_inventory")
try:
    if cfg:
        if st.sidebar.button("Actualiser maintenant", width="stretch"): load_google.clear()
        matrices, synced = load_google(*cfg); source = f"Google Sheets · {synced}"; st.sidebar.success("Google Sheets connecté")
    else: matrices=load_demo(); source="Classeur de démonstration local"; st.sidebar.warning("Google Sheets non configuré")
    sites, stocks = parse(matrices)
except Exception as exc:
    st.error(f"Chargement impossible : {exc}"); st.stop()

with st.sidebar:
    st.caption("Source : "+source); st.subheader("Filtres")
    zones=sorted(sites.zone.unique()); z=st.multiselect("Zone de santé",zones,default=zones)
    sites=sites[sites.zone.isin(z)]; stocks=stocks[stocks.zone.isin(z)]
    areas=sorted(sites.aire_sante.unique()); a=st.multiselect("Aire de santé",areas,default=areas)
    sites=sites[sites.aire_sante.isin(a)]; stocks=stocks[stocks.aire_sante.isin(a)]
    vaccines=sorted(stocks.vaccin.unique()); v=st.multiselect("Vaccin",vaccines,default=vaccines); stocks=stocks[stocks.vaccin.isin(v)]

if sites.empty: st.warning("Aucune donnée pour les filtres sélectionnés."); st.stop()
fridges=sites.refrigerateurs.fillna(0).sum(); known=sites.refrigerateur_fonctionnel|sites.refrigerateur_en_panne
functional=sites.loc[sites.refrigerateur_fonctionnel,"refrigerateurs"].fillna(0).sum(); functionality=functional/fridges if fridges else np.nan
tag_rate=sites.fridge_tag.mean(); maintenance=sites.maintenance_preventive.mean(); stockouts=int(stocks.rupture.sum())
stock_observed=stocks.stock_disponible.notna().sum(); stockout_rate=stocks.rupture.sum()/stock_observed if stock_observed else np.nan
low_stock_rate=(stocks.loc[stocks.couverture_semaines.notna(),"couverture_semaines"]<1).mean() if stocks.couverture_semaines.notna().any() else np.nan
fridge_reporting=sites.refrigerateurs.notna().mean(); status_reporting=(sites.statut_refrigerateur!="Non renseigné").mean()
planned_sessions=sites.seances_planifiees.sum(); completed_sessions=sites.seances_realisees.sum(); session_rate=completed_sessions/planned_sessions if planned_sessions else np.nan
st.markdown(f'''<div class="cards">
<div class="card" style="--accent:#0891b2"><div class="label">Réfrigérateurs recensés</div><div class="value">{integer(fridges)}</div><div class="hint">équipements déclarés</div></div>
<div class="card" style="--accent:#059669"><div class="label">Fonctionnalité</div><div class="value">{pct(functionality)}</div><div class="hint">équipements fonctionnels</div></div>
<div class="card" style="--accent:#dc2626"><div class="label">Équipements en panne</div><div class="value">{integer(sites.loc[sites.refrigerateur_en_panne,'refrigerateurs'].fillna(0).sum())}</div><div class="hint">action corrective requise</div></div>
<div class="card" style="--accent:#2563eb"><div class="label">Fridge-tag disponible</div><div class="value">{pct(tag_rate)}</div><div class="hint">aires équipées</div></div>
<div class="card" style="--accent:#d97706"><div class="label">Maintenance préventive</div><div class="value">{pct(maintenance)}</div><div class="hint">aires couvertes</div></div>
<div class="card" style="--accent:#e11d48"><div class="label">Ruptures de stock</div><div class="value">{stockouts}</div><div class="hint">vaccin × aire de santé</div></div></div>''',unsafe_allow_html=True)

module=st.radio("Module",["Synthèse décisionnelle","Vue générale","Stocks vaccins","Équipements","Maintenance et alertes","Activités","Données"],horizontal=True,label_visibility="collapsed")
notes={"Synthèse décisionnelle":"Lecture rapide des indicateurs phares, priorités et actions recommandées.","Vue générale":"Situation consolidée de la chaîne du froid dans les zones sélectionnées.","Stocks vaccins":"Besoins, stocks disponibles, semaines de couverture et ruptures par antigène.","Équipements":"Répartition et état fonctionnel des réfrigérateurs par zone et aire de santé.","Maintenance et alertes":"Liste opérationnelle des équipements et stocks nécessitant une action.","Activités":"Mise en relation des séances, supervisions et disponibilité des équipements.","Données":"Tables détaillées prêtes pour contrôle et téléchargement."}
st.markdown(f'<div class="module-note">{notes[module]}</div>',unsafe_allow_html=True)

if module=="Synthèse décisionnelle":
    critical_count=int(sites.refrigerateur_en_panne.sum()+stocks.rupture.sum())
    quality_ok=fridge_reporting>=.9 and status_reporting>=.9 and stock_observed/max(len(stocks),1)>=.9
    overall_level="Critique" if critical_count>0 else ("À surveiller" if maintenance<.8 or tag_rate<.8 or not quality_ok else "Satisfaisant")
    level_class="critical" if overall_level=="Critique" else ("watch" if overall_level=="À surveiller" else "good")
    st.markdown(f'''<div class="decision-grid">
    <div class="decision"><span class="badge {level_class}">{overall_level}</span><h4>Situation opérationnelle</h4><p>{critical_count} situation(s) critique(s) détectée(s), incluant les pannes et ruptures déclarées.</p></div>
    <div class="decision"><span class="badge {'good' if quality_ok else 'high'}">{'Données suffisantes' if quality_ok else 'Données à compléter'}</span><h4>Fiabilité de la lecture</h4><p>{pct(fridge_reporting)} des aires ont renseigné le nombre d’équipements et {pct(stock_observed/max(len(stocks),1))} les stocks.</p></div>
    <div class="decision"><span class="badge {'good' if maintenance>=.8 else 'watch'}">{pct(maintenance)}</span><h4>Prévention des pannes</h4><p>Part des aires déclarant une maintenance préventive. Cible indicative affichée : au moins 80 %.</p></div>
    </div>''',unsafe_allow_html=True)

    st.subheader("Informations importantes")
    insights=[]
    if stockouts:
        top=stocks[stocks.rupture].groupby("vaccin").size().sort_values(ascending=False)
        insights.append(("#dc2626","Ruptures de stock",f"{stockouts} rupture(s) déclarée(s). L’antigène le plus touché est {escape(str(top.index[0]))} dans {int(top.iloc[0])} aire(s)."))
    elif stock_observed: insights.append(("#059669","Stocks","Aucune rupture n’est déclarée dans les données actuellement filtrées."))
    else: insights.append(("#d97706","Stocks non renseignés","Aucun stock exploitable n’est disponible ; la situation des ruptures ne peut pas être confirmée."))
    if fridges==0: insights.append(("#d97706","Inventaire des équipements incomplet","Aucun réfrigérateur n’est recensé. Compléter les colonnes de nombre et d’état avant toute décision."))
    elif not pd.isna(functionality) and functionality<.8: insights.append(("#dc2626","Fonctionnalité insuffisante",f"Le taux de fonctionnalité est de {pct(functionality)}, sous la cible indicative de 80 %."))
    if tag_rate<.8: insights.append(("#2563eb","Suivi de température",f"Seulement {pct(tag_rate)} des aires déclarent un Fridge-tag disponible."))
    if maintenance<.8: insights.append(("#d97706","Maintenance préventive",f"La couverture de maintenance est de {pct(maintenance)} ; prioriser les aires non couvertes."))
    if not pd.isna(session_rate) and session_rate<.8: insights.append(("#7c3aed","Continuité des services",f"{pct(session_rate)} des séances planifiées sont déclarées réalisées."))
    for color,title,message in insights[:6]: st.markdown(f'<div class="insight" style="--accent:{color}"><strong>{title}</strong><br><span>{message}</span></div>',unsafe_allow_html=True)

    zone_sites=sites.groupby("zone",as_index=False).agg(aires=("aire_sante","nunique"),maintenance=("maintenance_preventive","mean"),fridge_tag=("fridge_tag","mean"),pannes=("refrigerateur_en_panne","sum"))
    zone_stock=stocks.groupby("zone",as_index=False).agg(ruptures=("rupture","sum"),couverture=("couverture_semaines","mean"))
    priorities=zone_sites.merge(zone_stock,on="zone",how="left").fillna({"ruptures":0,"pannes":0})
    priorities["score_priorite"]=np.minimum(priorities.ruptures*8,40)+np.minimum(priorities.pannes*20,30)+(1-priorities.maintenance)*20+(1-priorities.fridge_tag)*10
    priorities["priorite"]=pd.cut(priorities.score_priorite,[-1,15,30,45,np.inf],labels=["Faible","Modérée","Élevée","Critique"])
    priorities["maintenance_pct"]=priorities.maintenance*100; priorities["fridge_tag_pct"]=priorities.fridge_tag*100
    priorities=priorities.sort_values("score_priorite",ascending=False)
    c1,c2=st.columns([1.35,1])
    fig=px.bar(priorities,x="zone",y="score_priorite",color="priorite",title="Priorisation opérationnelle par zone",color_discrete_map={"Faible":"#16a34a","Modérée":"#eab308","Élevée":"#f97316","Critique":"#dc2626"})
    c1.plotly_chart(style(fig,360),width="stretch")
    c2.dataframe(priorities[["zone","priorite","score_priorite","ruptures","pannes","maintenance_pct","fridge_tag_pct"]],width="stretch",height=360,hide_index=True,column_config={"score_priorite":st.column_config.ProgressColumn("Score",min_value=0,max_value=100,format="%.0f"),"maintenance_pct":st.column_config.NumberColumn("Maintenance",format="%.0f%%"),"fridge_tag_pct":st.column_config.NumberColumn("Fridge-tag",format="%.0f%%")})
    st.markdown('<div class="legend-box"><b>Règles de lecture indicatives :</b> rupture ou panne = action immédiate ; couverture de stock &lt; 1 semaine = risque élevé ; fonctionnalité, maintenance ou Fridge-tag &lt; 80 % = attention. Les seuils doivent être validés par le PEV.</div>',unsafe_allow_html=True)
elif module=="Vue générale":
    by_zone=sites.groupby("zone",as_index=False).agg(population=("population","sum"),refrigerateurs=("refrigerateurs","sum"),aires=("aire_sante","nunique"),fridge_tag=("fridge_tag","mean"),maintenance=("maintenance_preventive","mean"))
    rupt=stocks.groupby("zone",as_index=False).agg(ruptures=("rupture","sum")); by_zone=by_zone.merge(rupt,on="zone",how="left")
    c1,c2=st.columns(2)
    c1.plotly_chart(style(px.bar(by_zone,x="zone",y="refrigerateurs",color="ruptures",title="Équipements et ruptures par zone",color_continuous_scale=["#0ea5a8","#f59e0b","#dc2626"])),width="stretch")
    heat=stocks.pivot_table(index="vaccin",columns="zone",values="couverture_semaines",aggfunc="mean")
    c2.plotly_chart(style(px.imshow(heat,text_auto=".1f",aspect="auto",title="Semaines moyennes de couverture",color_continuous_scale=["#dc2626","#facc15","#059669"])),width="stretch")
    st.dataframe(by_zone,width="stretch",hide_index=True)
elif module=="Stocks vaccins":
    vac=stocks.groupby("vaccin",as_index=False).agg(besoin_hebdomadaire=("besoin_hebdomadaire","sum"),stock_disponible=("stock_disponible","sum"),ruptures=("rupture","sum")); vac["couverture_semaines"]=vac.stock_disponible/vac.besoin_hebdomadaire.replace(0,np.nan)
    fig=px.bar(vac,x="vaccin",y="couverture_semaines",color="couverture_semaines",title="Couverture de stock par vaccin",color_continuous_scale=["#dc2626","#facc15","#059669"]); fig.add_hline(y=1,line_dash="dash",line_color="#dc2626",annotation_text="Seuil 1 semaine")
    st.plotly_chart(style(fig),width="stretch")
    risk=stocks[(stocks.rupture)|(stocks.couverture_semaines<1)].sort_values(["rupture","couverture_semaines"],ascending=[False,True]); st.dataframe(risk,width="stretch",hide_index=True)
    st.download_button("Télécharger l’état des stocks",csv_bytes(stocks),"stocks_vaccins.csv","text/csv")
elif module=="Équipements":
    status=sites.groupby(["zone","statut_refrigerateur"],as_index=False)["refrigerateurs"].sum()
    st.plotly_chart(style(px.bar(status,x="zone",y="refrigerateurs",color="statut_refrigerateur",title="État des réfrigérateurs par zone",barmode="stack")),width="stretch")
    st.dataframe(sites[["zone","aire_sante","refrigerateurs","statut_refrigerateur","fridge_tag","date_debut_panne"]],width="stretch",hide_index=True)
elif module=="Maintenance et alertes":
    alerts=[]
    for _,r in sites.iterrows():
        if r.refrigerateur_en_panne: alerts.append(["Critique",r.zone,r.aire_sante,"Réfrigérateur en panne",r.date_debut_panne])
        if not r.fridge_tag: alerts.append(["Élevée",r.zone,r.aire_sante,"Fridge-tag absent ou non déclaré",""])
        if not r.maintenance_preventive: alerts.append(["Élevée",r.zone,r.aire_sante,"Maintenance préventive non réalisée",""])
    for _,r in stocks[stocks.rupture].iterrows(): alerts.append(["Critique",r.zone,r.aire_sante,f"Rupture de {r.vaccin}",""])
    alert_df=pd.DataFrame(alerts,columns=["Niveau","Zone","Aire de santé","Alerte","Date début panne"])
    if alert_df.empty: st.success("Aucune alerte active.")
    else:
        st.error(f"{(alert_df.Niveau=='Critique').sum()} alertes critiques et {(alert_df.Niveau=='Élevée').sum()} alertes élevées.")
        st.dataframe(alert_df,width="stretch",hide_index=True); st.download_button("Télécharger les alertes",csv_bytes(alert_df),"alertes_chaine_froid.csv","text/csv")
elif module=="Activités":
    act=sites.groupby("zone",as_index=False).agg(seances_planifiees=("seances_planifiees","sum"),seances_realisees=("seances_realisees","sum"),supervisions_prevues=("supervisions_prevues","sum"),supervisions_realisees=("supervisions_realisees","sum")); act["taux_seances"]=act.seances_realisees/act.seances_planifiees.replace(0,np.nan); act["taux_supervision"]=act.supervisions_realisees/act.supervisions_prevues.replace(0,np.nan)
    chart=act.melt("zone",value_vars=["taux_seances","taux_supervision"],var_name="Indicateur",value_name="Taux"); fig=px.bar(chart,x="zone",y="Taux",color="Indicateur",barmode="group",title="Réalisation des activités"); fig.update_yaxes(tickformat=".0%")
    st.plotly_chart(style(fig),width="stretch"); st.dataframe(act,width="stretch",hide_index=True)
else:
    tab1,tab2=st.tabs(["Équipements et activités","Stocks"])
    tab1.dataframe(sites,width="stretch",hide_index=True); tab1.download_button("Télécharger les équipements",csv_bytes(sites),"equipements_chaine_froid.csv","text/csv")
    tab2.dataframe(stocks,width="stretch",hide_index=True); tab2.download_button("Télécharger les stocks",csv_bytes(stocks),"stocks_chaine_froid.csv","text/csv")

st.caption("Les seuils et classifications doivent être validés selon les normes nationales du PEV avant utilisation opérationnelle.")
