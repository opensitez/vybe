# vybe-test: python/try_finally_flow/try_finally_with_break_in_finally_not_allowed_use_pattern
# origin: languages/python/tests/python/test_try_finally_flow.rs

out = []
for i in range(2):
 try:
  out.append(i)
 finally:
  if i == 1:
   pass
print(out)
