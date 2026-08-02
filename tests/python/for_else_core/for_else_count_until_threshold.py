# vybe-test: python/for_else_core/for_else_count_until_threshold
# origin: languages/python/tests/python/test_for_else_core.rs

count = 0
for _ in range(10):
 count += 1
 if count == 3:
  break
else:
 print('full')
print(count)
