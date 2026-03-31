dm = {"a": 1, "b": 2}
s = 0
for k in dm:
    s = s + dm[k]
if s == 3:
    print("PASS test_for_dict")
else:
    print("FAIL test_for_dict")
