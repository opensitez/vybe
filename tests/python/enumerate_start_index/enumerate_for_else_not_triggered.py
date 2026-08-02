# vybe-test: python/enumerate_start_index/enumerate_for_else_not_triggered
# origin: languages/python/tests/python/test_enumerate_start_index.rs

for i, v in enumerate([1]):
 pass
else:
 print(i)
