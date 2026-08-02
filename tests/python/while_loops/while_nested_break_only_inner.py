# vybe-test: python/while_loops/while_nested_break_only_inner
# origin: languages/python/tests/python/test_while_loops.rs

out = []
i = 0
while i < 3:
 j = 0
 while j < 3:
  if j == 1:
   break
  out.append(j)
  j += 1
 i += 1
print(len(out))
