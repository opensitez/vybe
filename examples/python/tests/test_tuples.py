# Focused tuple tests: creation and indexing
tp = (4, 5, 6)
ok = True
if not (tp[0] == 4 and tp[2] == 6): ok = False
if ok:
  print("PASS test_tuples")
else:
  print("FAIL test_tuples")
