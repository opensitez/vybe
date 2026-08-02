# vybe-test: python/try_finally_flow/finally_runs_on_continue_in_loop
# origin: languages/python/tests/python/test_try_finally_flow.rs

out = []
for i in range(2):
 try:
  if i == 0:
   continue
  out.append(i)
 finally:
  out.append(9)
print(out)
