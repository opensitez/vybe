# vybe-test: python/try_finally_flow/try_nested_finally_only_inner_on_inner_break
# origin: languages/python/tests/python/test_try_finally_flow.rs

out = []
for _ in range(1):
 try:
  try:
   break
  finally:
   out.append('i')
 finally:
  out.append('o')
print(out)
