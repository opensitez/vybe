# vybe-test: python/nested_loop_control/break_in_list_iteration_nested_in_dict_keys
# origin: languages/python/tests/python/test_nested_loop_control.rs

d = {'a': [1, 2], 'b': [3]}
out = []
for k in d:
 for v in d[k]:
  if v == 2:
   break
  out.append(k + str(v))
print(out)
