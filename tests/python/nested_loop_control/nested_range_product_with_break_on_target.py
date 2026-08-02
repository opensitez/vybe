# vybe-test: python/nested_loop_control/nested_range_product_with_break_on_target
# origin: languages/python/tests/python/test_nested_loop_control.rs

target = 5
found = False
for a in range(3):
 for b in range(3):
  if a * 3 + b == target:
   found = True
   break
 if found:
  break
print(found)
