# vybe-test: python/for_else_core/for_else_count_completes_full_range
# origin: languages/python/tests/python/test_for_else_core.rs

count = 0
for _ in range(3):
 count += 1
else:
 print('full')
print(count)
