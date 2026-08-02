# vybe-test: python/try_finally_flow/try_with_continue_in_loop
# origin: languages/python/tests/python/test_try_finally_flow.rs

out = []
for i in range(3):
 try:
  if i == 1:
   continue
  out.append(i)
 finally:
  pass
print(out)
