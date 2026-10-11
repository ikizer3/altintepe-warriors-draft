"""Offline invariants for the ESPN sync; no league data is modified."""
import unittest

SCORING = {
 "FGM":1,"FGA":-0.75,"FTM":1,"FTA":-1,"3PM":0.8,
 "OREB":0.5,"REB":0.8,"AST":1.5,"STL":2.25,"BLK":2,
 "TO":-1.5,"EJ":-5,"FF":-3,"PF":-0.2,"TF":-1,
 "DQ":-3,"DD":3,"TD":6,"QD":20,"PTS":1,"TW":1,
}

def score(stats):
    unknown=set(stats)-set(SCORING)
    if unknown:
        raise ValueError("Unknown stats: "+str(sorted(unknown)))
    return round(sum(float(stats.get(k,0))*v for k,v in SCORING.items()),5)

def validate(rosters,available):
    ids=[str(p) for team in rosters.values() for p in team]
    free=[str(p) for p in available]
    if len(ids)!=len(set(ids)):raise ValueError("Duplicate rostered player")
    if len(free)!=len(set(free)):raise ValueError("Duplicate available player")
    if set(ids)&set(free):raise ValueError("Roster/FA overlap")
    return len(ids)+len(free)

def roster_changes(before,after):
    changes=[]
    for team_id in sorted(set(before)|set(after)):
        old=set(before.get(team_id,[]))
        new=set(after.get(team_id,[]))
        changes.extend((team_id,"ADD",pid) for pid in sorted(new-old))
        changes.extend((team_id,"DROP",pid) for pid in sorted(old-new))
    return changes

def injury_changes(before,after):
    return [(pid,before[pid],after[pid]) for pid in sorted(set(before)&set(after))
            if before[pid]!=after[pid]]

class SyncTests(unittest.TestCase):
    def test_master_scoring_21_categories(self):
        self.assertEqual(len(SCORING),21)
        self.assertEqual(score({"PTS":20,"FGM":8,"FGA":15,"FTM":4,"FTA":5,"3PM":3,"REB":7,"OREB":2,"AST":6,"STL":2,"BLK":1,"TO":3,"PF":2,"TW":1}),36.35)
    def test_scoring_negative_and_bonus(self):
        self.assertEqual(score({"EJ":1,"FF":1,"TF":1,"DQ":1,"DD":1,"TD":1,"QD":1}),17)
    def test_duplicate_rostered(self):
        with self.assertRaises(ValueError):validate({"a":[1],"b":[1]},[2])
    def test_roster_available_overlap(self):
        with self.assertRaises(ValueError):validate({"a":[1,2]},[2,3])
    def test_duplicate_available(self):
        with self.assertRaises(ValueError):validate({"a":[1]},[2,2])
    def test_valid_transfer(self):
        self.assertEqual(validate({"a":[1],"b":[2]},[3,4]),4)
    def test_add_drop_between_teams(self):
        self.assertEqual(roster_changes({"a":[1,2],"b":[3]},{"a":[1],"b":[2,3]}),
                         [("a","DROP",2),("b","ADD",2)])
    def test_waiver_drop(self):
        self.assertEqual(roster_changes({"a":[1,2]},{"a":[1]}),[("a","DROP",2)])
        self.assertEqual(validate({"a":[1]},[2,3]),3)
    def test_injury_transition(self):
        self.assertEqual(injury_changes({1:"ACTIVE",2:"DTD"},{1:"OUT",2:"DTD"}),[(1,"ACTIVE","OUT")])
    def test_injury_addition_not_a_change(self):
        self.assertEqual(injury_changes({1:"OUT"},{1:"ACTIVE",2:"OUT"}),[(1,"OUT","ACTIVE")])
    def test_unknown_stat_fails(self):
        with self.assertRaises(ValueError):score({"UNKNOWN":2})

if __name__=="__main__":
    unittest.main()
