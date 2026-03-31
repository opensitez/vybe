s = 0
for i in [1,2,3,4]:
    if i == 3:
        break
    s = s + i
if s == 3:
    print("PASS test_for_break")
else:
    print("FAIL test_for_break")

s = 0
for i in [1,2,3,4]:
    if i % 2 == 0:
        continue
    s = s + i
if s == 1 + 3:
    print("PASS test_for_continue")
else:
    print("FAIL test_for_continue")
