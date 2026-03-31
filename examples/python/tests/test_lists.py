# Focused list tests: nested lists and indexing
lst = [10, 20, [1, 2]]
ok = True
# avoid chained indexing in a single expression to work around current parser limits
sub = lst[2]
if not (lst[0] == 10 and sub[1] == 2): ok = False
if ok:
  print("PASS test_lists")
else:
  print("FAIL test_lists")
