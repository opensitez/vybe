# vybe-test: python/for_while_extended/for_star_unpack
# origin: languages/python/tests/python/test_for_while_extended.rs

for h, *t in [(1, 2, 3)]:
 print(h, len(t))
