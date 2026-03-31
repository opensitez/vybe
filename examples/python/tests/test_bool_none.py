# Test booleans, None, and simple ops
flag = True
nothing = None
sum0 = 1 + 2
prod0 = 2 * 3

ok = True
if not flag: ok = False
if not (nothing == None): ok = False
if not (sum0 == 3 and prod0 == 6): ok = False

if ok:
  print("PASS test_bool_none")
else:
  print("FAIL test_bool_none")
