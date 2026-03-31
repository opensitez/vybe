# Focused dict tests: literal creation and string-key lookups
dm = {"x": 7, "y": 8}
ok = True
if not (dm["x"] == 7 and dm["y"] == 8): ok = False
if ok:
  print("PASS test_dicts")
else:
  print("FAIL test_dicts")
