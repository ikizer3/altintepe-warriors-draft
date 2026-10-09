"""Hourly ESPN fantasy league archive; never edits the draft master workbook."""
import datetime, json, pathlib, urllib.request, urllib.parse, urllib.error
from openpyxl import Workbook, load_workbook
ROOT=pathlib.Path(__file__).resolve().parents[1]
DIR=ROOT/"league_archive"
DIR.mkdir(exist_ok=True)
BASE="https://lm-api-reads.fantasy.espn.com/apis/v3/games/fba/seasons/2027/segments/0/leagues/88948640"
BOOK=DIR/"ALTINTEPE_WARRIORS_LIVE_LEAGUE_2026-27.xlsx"
NOW=datetime.datetime.now(datetime.timezone.utc).isoformat()
def fetch(views,filters=None,params=None):
    q=[("view",v) for v in views]
    q+=list((params or {}).items())
    url=BASE+"?"+urllib.parse.urlencode(q)
    headers={"User-Agent":"Mozilla/5.0","Accept":"application/json"}
    if filters:headers["x-fantasy-filter"]=json.dumps(filters,separators=(",",":"))
    req=urllib.request.Request(url,headers=headers)
    with urllib.request.urlopen(req,timeout=45) as resp:return json.load(resp)
def dump(path,obj):
    path.write_text(json.dumps(obj,ensure_ascii=False,indent=2,sort_keys=True,default=str)+"\n",encoding="utf-8")
def sheet(wb,name,headers,rows):
    ws=wb[name] if name in wb else wb.create_sheet(name)
    ws.delete_rows(1,ws.max_row)
    ws.append(headers)
    for row in rows:ws.append([v if isinstance(v,(str,int,float,bool,type(None))) else json.dumps(v,ensure_ascii=False) for v in row])
    ws.freeze_panes="A2"
    ws.auto_filter.ref=ws.dimensions
    for col,width in {"A":23,"B":24,"C":28,"D":22,"E":30,"F":22,"G":23}.items():ws.column_dimensions[col].width=width
    return ws
def main():
    data=fetch(["mTeam","mRoster","mStatus","mMatchup","mTransactions2"])
    teams=data.get("teams",[])
    if data.get("id")!=88948640 or len(teams)!=16:raise RuntimeError("Unexpected league or incomplete team list")
    current={}
    for t in teams:
        tid=str(t["id"])
        entries=(t.get("roster") or {}).get("entries")
        if not entries:raise RuntimeError("Missing roster for team "+tid)
        name=t.get("name") or (str(t.get("location") or "")+" "+str(t.get("nickname") or "")).strip()
        current[tid]={"name":name,"players":[]}
        for e in entries:
            p=((e.get("playerPoolEntry") or {}).get("player") or {})
            current[tid]["players"].append({"id":e.get("playerId"),"name":p.get("fullName"),"slot":e.get("lineupSlotId"),"injury":p.get("injuryStatus"),"stats":p.get("stats") or []})
    oldfile=DIR/"latest.json"
    old=json.loads(oldfile.read_text(encoding="utf-8")) if oldfile.exists() else {}
    changes=[]
    for tid,t in current.items():
        before={str(p["id"]):p for p in old.get("teams",{}).get(tid,{}).get("players",[])}
        after={str(p["id"]):p for p in t["players"]}
        if old:
            for pid in sorted(after.keys()-before.keys()):changes.append([NOW,t["name"],"ADD",after[pid]["name"],pid])
            for pid in sorted(before.keys()-after.keys()):changes.append([NOW,t["name"],"DROP",before[pid]["name"],pid])
    # ESPN transaction IDs prevent duplicates and preserve movements between hourly polls.
    tx=data.get("transactions") or []
    if not isinstance(tx,list):tx=[]
    txfile=DIR/"transactions.json"
    previous=json.loads(txfile.read_text(encoding="utf-8")) if txfile.exists() else {}
    txmap=dict(previous)
    for x in tx:
        key=str(x.get("id") or x.get("transactionId") or "")
        if key:txmap[key]=x
    # Free agent endpoint is best-effort. Do not erase last good FA snapshot on failure.
    fa_path=DIR/"free_agents.json"
    fa_error=""
    try:
        filt={"players":{"filterStatus":{"value":["FREEAGENT","WAIVERS"]},"limit":1500,"sortPercOwned":{"sortPriority":1,"sortAsc":False}}}
        response=fetch(["kona_player_info"],filt)
        players=response.get("players") or []
        if not isinstance(players,list):raise ValueError("Unexpected player payload")
        free=[]
        for x in players:
            p=x.get("player") or {}
            free.append({"id":p.get("id"),"name":p.get("fullName"),"status":x.get("status"),"injury":p.get("injuryStatus"),"ownership":(p.get("ownership") or {}).get("percentOwned")})
        if not free:raise ValueError("Empty FA result; previous data retained")
        roster_ids={str(p["id"]) for team in current.values() for p in team["players"]}
        free=[p for p in free if str(p.get("id")) not in roster_ids]
        if not free:raise ValueError("No unrostered players after roster exclusion")
        dump(fa_path,{"updatedAt":NOW,"players":free,"rosteredExcluded":len(roster_ids)})
    except Exception as exc:fa_error=str(exc)[:350]
    # Preserve a snapshot and change history; this is separate from the master file.
    latest={"updatedAt":NOW,"leagueId":88948640,"teams":current,"status":data.get("status"),"schedule":data.get("schedule")}
    dump(oldfile,latest)
    dump(txfile,txmap)
    dump(DIR/("snapshot-"+NOW[:10]+".json"),latest)
    if changes:
        with (DIR/"events.jsonl").open("a",encoding="utf-8") as fh:
            for event in changes:fh.write(json.dumps(event,ensure_ascii=False)+"\n")
    wb=load_workbook(BOOK) if BOOK.exists() else Workbook()
    sheet(wb,"CANLI_KADROLAR",["Kontrol UTC","Takım ID","Takım","Oyuncu ID","Oyuncu","Kadro Slotu","Sakatlık"],[
        [NOW,int(tid),t["name"],p["id"],p["name"],p["slot"],p["injury"]]
        for tid,t in sorted(current.items(),key=lambda item:int(item[0])) for p in t["players"]])
    if "Sheet" in wb and len(wb.sheetnames)>1:del wb["Sheet"]
    events=[]
    evfile=DIR/"events.jsonl"
    if evfile.exists():
        for line in evfile.read_text(encoding="utf-8").splitlines():
            if line.strip():events.append(json.loads(line))
    sheet(wb,"KADRO_DEGISIMLERI",["Gözlem UTC","Takım","İşlem","Oyuncu","ESPN ID"],events)
    txrows=[]
    for key,t in sorted(txmap.items(),key=lambda x:str(x[0])):
        txrows.append([key,t.get("type"),t.get("status"),t.get("processDate"),json.dumps(t.get("items") or [],ensure_ascii=False),json.dumps(t,ensure_ascii=False)])
    sheet(wb,"ESPN_ISLEMLERI",["İşlem ID","Tür","Durum","Tarih","Oyuncu İşlemleri","Ham ESPN Kaydı"],txrows)
    if fa_path.exists():
        fa=json.loads(fa_path.read_text(encoding="utf-8"))
        roster_ids={str(p["id"]) for team in current.values() for p in team["players"]}
        fa["players"]=[p for p in fa.get("players",[]) if str(p.get("id")) not in roster_ids]
        sheet(wb,"FA_WAIVER",["Son başarılı FA kontrolü","Oyuncu ID","Oyuncu","ESPN Durumu","Sakatlık","Sahiplik %"],[
            [fa.get("updatedAt"),p.get("id"),p.get("name"),p.get("status"),p.get("injury"),p.get("ownership")] for p in fa.get("players",[])])
    schedule=data.get("schedule") or []
    sheet(wb,"ESLESMELER",["Matchup ID","Periyot","Ev Sahibi ID","Deplasman ID","Ev Sahibi Puan","Deplasman Puan"],[
        [s.get("id"),s.get("matchupPeriodId"),(s.get("home") or {}).get("teamId"),(s.get("away") or {}).get("teamId"),(s.get("home") or {}).get("totalPoints"),(s.get("away") or {}).get("totalPoints")] for s in schedule if isinstance(s,dict)])
    # ESPN player.stats entries can represent projected, season-total or period data.
    # Archive them raw; do not mislabel projections as actual daily scores.
    stats=[]
    for tid,t in current.items():
        for p in t["players"]:
            for s in p["stats"]:
                stats.append([NOW,t["name"],p["id"],p["name"],s.get("scoringPeriodId"),s.get("statSourceId"),s.get("appliedTotal"),json.dumps(s,ensure_ascii=False)])
    sheet(wb,"ESPN_OYUNCU_STATS",["Kontrol UTC","Takım","Oyuncu ID","Oyuncu","Skor Periyodu","Kaynak (0 gerçek, 1 tahmin)","ESPN Uygulanan Puan","Ham Stat"],stats)
    sheet(wb,"SENKRON_DURUMU",["Alan","Değer"],[
        ["Son başarılı kadro kontrolü (UTC)",NOW],["Lig ID",88948640],["Takım sayısı",len(teams)],
        ["Kadro değişikliği",len(changes)],["ESPN işlem sayısı",len(txmap)],
        ["FA son başarılı kontrol",json.loads(fa_path.read_text()).get("updatedAt") if fa_path.exists() else "YOK"],
        ["FA hatası",fa_error or "YOK"],["Not","İşlem geçmişi ESPN API'nin sunduğu kapsamla sınırlıdır."]])
    wb.save(BOOK)
    print("OK teams=",len(teams),"roster changes=",len(changes),"transactions=",len(txmap),"FA error=",fa_error)
if __name__=="__main__":main()
