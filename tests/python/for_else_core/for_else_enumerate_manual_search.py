# vybe-test: python/for_else_core/for_else_enumerate_manual_search
# origin: languages/python/tests/python/test_for_else_core.rs

words = ['cat', 'dog', 'bird']
for i in range(len(words)):
 if words[i] == 'dog':
  print(i)
  break
else:
 print(-1)
