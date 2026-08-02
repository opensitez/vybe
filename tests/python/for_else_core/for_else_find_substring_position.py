# vybe-test: python/for_else_core/for_else_find_substring_position
# origin: languages/python/tests/python/test_for_else_core.rs

hay = 'hello'
needle = 'll'
for i in range(len(hay) - len(needle) + 1):
 if hay[i:i+len(needle)] == needle:
  print(i)
  break
else:
 print(-1)
