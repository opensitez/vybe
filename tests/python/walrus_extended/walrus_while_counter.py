# vybe-test: python/walrus_extended/walrus_while_counter
# origin: languages/python/tests/python/test_walrus_extended.rs

n = 3
c = 0
while (n := n - 1) >= 0:
 c += 1
print(c)
