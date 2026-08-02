# vybe-test: python/nested_loop_control/break_from_inner_for_on_string_chars
# origin: languages/python/tests/python/test_nested_loop_control.rs

out = []
for ch in 'abc':
 for _ in range(3):
  if ch == 'b':
   break
  out.append(ch)
print(out)
