# Test list/tuple/dict literals and indexing
lst = [1, 2, "three"]
tp = (1, 2, 3)
dm = {"a": 1, "b": 2}

ok = True
if not (lst[0] == 1 and lst[2] == "three"): ok = False
if not (tp[1] == 2): ok = False
if not (dm["b"] == 2): ok = False

if ok:
  print("PASS test_literals")
else:
  print("FAIL test_literals")
