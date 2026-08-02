# vybe-test: python/while_loops/while_membership_scan_finds_target
# origin: languages/python/tests/python/test_while_loops.rs

xs = [4, 7, 9]
i = 0
found = False
while i < len(xs) and not found:
 found = xs[i] == 7
 i += 1
print(found)
