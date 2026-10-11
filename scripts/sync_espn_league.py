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
    # Split ESPN views to avoid oversized/partial combined responses.
    data=fetch(["mTeam","mRoster","mStatus"])
    matchup=fetch(["mMatchupScore"])
    if not isinstance(matchup.get("schedule"),list):
        raise RuntimeError("Missing ESPN matchup schedule")
    data["schedule"]=matchup["schedule"]
    # mTransactions2 is scoring-period scoped; an unscoped response may be partial.
    period=int((data.get("status") or {}).get("currentScoringPeriod") or 1)
    if not 1 <= period <= 400:
        raise RuntimeError("Invalid ESPN scoring period: "+str(period))
    tx_by_id={}
    for scoring_period in range(1,period+1):
        batch=fetch(["mTransactions2"],params={"scoringPeriodId":scoring_period})
        items=batch.get("transactions")
        if not isinstance(items,list):
            raise RuntimeError("Missing transactions in scoring period "+str(scoring_period))
        for item in items:
            key=str(item.get("id") or item.get("transactionId") or "")
            if key:tx_by_id[key]=item
    data["transactions"]=list(tx_by_id.values())
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
    if not isinstance(previous,dict):raise RuntimeError("Invalid previous transaction archive")
    txmap=dict(previous)
    for x in tx:
        key=str(x.get("id") or x.get("transactionId") or "")
        if key:txmap[key]=x
    # Free agent endpoint is best-effort. Do not erase last good FA snapshot on failure.
    fa_path=DIR/"free_agents.json"
    fa_error=""
    try:
        roster_ids={str(p["id"]) for team in current.values() for p in team["players"]}
        by_id={}
        counts={}
        # Query each ESPN availability category independently, rather than
        # trusting a single combined query or treating all unrostered as waivers.
        for requested_status in ("FREEAGENT","WAIVERS"):
            filt={"players":{"filterStatus":{"value":[requested_status]},"limit":2500,"sortPercOwned":{"sortPriority":1,"sortAsc":False}}}
            response=fetch(["kona_player_info"],filt)
            players=response.get("players")
            if not isinstance(players,list):raise ValueError("Invalid "+requested_status+" payload")
            counts[requested_status]=len(players)
            for x in players:
                p=x.get("player") or {}
                pid=p.get("id") or x.get("id")
                if pid is None or str(pid) in roster_ids:continue
                raw_status=x.get("status")
                # onTeamId is a second independent check; exclude all rostered.
                on_team=x.get("onTeamId") or 0
                if on_team:continue
                entry=by_id.setdefault(str(pid),{"id":pid,"name":p.get("fullName"),"status":None,"injury":p.get("injuryStatus"),"ownership":(p.get("ownership") or {}).get("percentOwned"),"waiverProcessDate":x.get("waiverProcessDate"),"rawStatus":raw_status,"matchedFilters":[]})
                entry["matchedFilters"].append(requested_status)
                if raw_status in ("FREEAGENT","WAIVERS"):
                    entry["status"]=raw_status
        free=list(by_id.values())
        if not free:raise ValueError("Empty available player pool; previous data retained")
        # Conflicting filters / labels must be surfaced, never silently guessed.
        for p in free:
            matched=set(p["matchedFilters"])
            if len(matched)>1 or p["status"] not in ("FREEAGENT","WAIVERS") or p["status"] not in matched:
                p["status"]="UNVERIFIED"
        dump(fa_path,{"updatedAt":NOW,"players":free,"rosteredExcluded":len(roster_ids),"filterCounts":counts,"unverifiedCount":sum(p["status"]=="UNVERIFIED" for p in free)})
    except Exception as exc:fa_error=str(exc)[:350]
    # Independent ESPN season-wide player list audit. Never infer completeness
    # from the league's available-player endpoint alone.
    audit={"checkedAt":NOW,"status":"UNVERIFIED","leagueRosterCount":sum(len(t["players"]) for t in current.values())}
    try:
        general_url="https://lm-api-reads.fantasy.espn.com/apis/v3/games/fba/seasons/2027/players?view=players_wl"
        general_filter=json.dumps({"filterActive":{"value":True}},separators=(",",":"))
        req=urllib.request.Request(general_url,headers={"User-Agent":"Mozilla/5.0","Accept":"application/json","x-fantasy-filter":general_filter})
        with urllib.request.urlopen(req,timeout=45) as resp:general=json.load(resp)
        if not isinstance(general,list) or not general:raise ValueError("Season-wide ESPN players list is empty or invalid")
        general_ids={str(p.get("id")) for p in general if isinstance(p,dict) and p.get("id") is not None}
        if not general_ids:raise ValueError("No player IDs in season-wide response")
        roster_ids={str(p["id"]) for t in current.values() for p in t["players"]}
        available=json.loads(fa_path.read_text(encoding="utf-8")) if fa_path.exists() else {}
        if available.get("updatedAt")!=NOW:raise ValueError("FA snapshot not refreshed this run")
        available_ids={str(p["id"]) for p in available.get("players",[]) if p.get("id") is not None}
        union=roster_ids|available_ids
        audit.update({"generalActiveCount":len(general_ids),"leagueUnionCount":len(union),"rosteredCount":len(roster_ids),"availableCount":len(available_ids),"generalNotInLeagueCount":len(general_ids-union),"leagueNotInGeneralActiveCount":len(union-general_ids),"generalNotInLeagueExamples":sorted(general_ids-union)[:30],"leagueNotInGeneralActiveExamples":sorted(union-general_ids)[:30],"status":"COMPARED"})
    except Exception as exc:
        audit["error"]=str(exc)[:350]
    if audit.get("status")!="COMPARED":
        raise RuntimeError("Independent ESPN player pool comparison unavailable: "+str(audit.get("error")))
    if audit["generalNotInLeagueCount"] or audit["leagueNotInGeneralActiveCount"]:
        raise RuntimeError("ESPN player pool mismatch: "+json.dumps(audit,ensure_ascii=False))
    dump(DIR/"player_pool_audit.json",audit)
    # Fail closed on inconsistent ESPN roster/FA/transaction data.
    roster_ids=[str(p["id"]) for t in current.values() for p in t["players"]]
    if len(roster_ids)!=len(set(roster_ids)):
        raise RuntimeError("Duplicate player across ESPN fantasy rosters")
    if fa_error:
        raise RuntimeError("FA/waiver refresh failed; refusing inconsistent snapshot: "+fa_error)
    available=json.loads(fa_path.read_text(encoding="utf-8"))
    available_ids=[str(p["id"]) for p in available.get("players",[])]
    if len(available_ids)!=len(set(available_ids)):
        raise RuntimeError("Duplicate player in FA/waiver pool")
    if set(roster_ids)&set(available_ids):
        raise RuntimeError("Rostered player appears in FA/waiver pool")
    if available.get("unverifiedCount"):
        raise RuntimeError("Unverified FA/waiver statuses; refusing snapshot")
    if len(txmap)<len(previous):
        raise RuntimeError("Transaction history shrank unexpectedly")
    # Validate roster movements against the resulting current roster.
    for change in changes:
        _,team_name,action,_,player_id=change
        matching=[t for t in current.values() if t["name"]==team_name]
        if len(matching)!=1:raise RuntimeError("Ambiguous team in roster change")
        now_ids={str(p["id"]) for p in matching[0]["players"]}
        if action=="ADD" and str(player_id) not in now_ids:
            raise RuntimeError("ADD missing from destination roster")
        if action=="DROP" and str(player_id) in now_ids:
            raise RuntimeError("DROP still in original roster")
    # Injury status changes are independently tracked across successive snapshots.
    injury_changes=[]
    for tid,team in current.items():
        prior={str(p["id"]):p for p in old.get("teams",{}).get(tid,{}).get("players",[])}
        for player in team["players"]:
            pid=str(player["id"])
            if pid in prior and prior[pid].get("injury")!=player.get("injury"):
                injury_changes.append({"checkedAt":NOW,"teamId":tid,"team":team["name"],"playerId":pid,"player":player["name"],"previous":prior[pid].get("injury"),"current":player.get("injury")})
    injury_path=DIR/"injury_changes.json"
    injury_history=json.loads(injury_path.read_text(encoding="utf-8")) if injury_path.exists() else []
    if not isinstance(injury_history,list):raise RuntimeError("Invalid injury change history")
    if injury_changes:dump(injury_path,injury_history+injury_changes)
    # Transactions are retained by unique ESPN ID; trade details remain in raw ESPN records.
    trade_records=[{"id":key,"type":v.get("type"),"status":v.get("status"),"processDate":v.get("processDate"),"items":v.get("items",[])} for key,v in txmap.items() if "TRADE" in str(v.get("type","")).upper()]
    dump(DIR/"trade_audit.json",{"checkedAt":NOW,"status":"OBSERVATION_ONLY","count":len(trade_records),"trades":trade_records})
    dump(DIR/"integrity_audit.json",{"checkedAt":NOW,"status":"PASS","teams":len(teams),"rostered":len(roster_ids),"available":len(available_ids),"transactions":len(txmap),"rosterChanges":len(changes),"overlap":0,"unverified":0})
    # Persist a machine-readable health report for external monitors.
    health={"checkedAt":NOW,"status":"HEALTHY","leagueId":88948640,
            "teams":len(teams),"rostered":len(roster_ids),"freeAgentsAndWaivers":len(available_ids),
            "transactions":len(txmap),"playerPoolMatch":True,"regressionTests":"GITHUB_ACTIONS_REQUIRED",
            "waiverPolicy":"ESPN_SOURCE_OF_TRUTH","alerts":[]}
    dump(DIR/"health.json",health)
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
        sheet(wb,"FA_WAIVER",["Son başarılı FA kontrolü","Oyuncu ID","Oyuncu","Doğrulanan durum","Sakatlık","Sahiplik %","Waiver bitiş tarihi","Ham ESPN durumu"],[
            [fa.get("updatedAt"),p.get("id"),p.get("name"),p.get("status"),p.get("injury"),p.get("ownership"),p.get("waiverProcessDate"),p.get("rawStatus")] for p in fa.get("players",[])])
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
        ["FA hatası",fa_error or "YOK"],["FA/WAIVER filtre sonuçları",json.dumps(json.loads(fa_path.read_text()).get("filterCounts",{})) if fa_path.exists() else "YOK"],["Belirsiz oyuncu",json.loads(fa_path.read_text()).get("unverifiedCount") if fa_path.exists() else "YOK"],["Not","İşlem geçmişi ESPN API'nin sunduğu kapsamla sınırlıdır."]])
    wb.save(BOOK)
    print("OK teams=",len(teams),"roster changes=",len(changes),"transactions=",len(txmap),"FA error=",fa_error)
if __name__=="__main__":main()
