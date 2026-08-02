# vybe-test: python/global_nonlocal/closure_capture_loop_variable_without_nonlocal_bug_pattern
# origin: languages/python/tests/python/test_global_nonlocal.rs

funcs = []
for i in range(2):
 funcs.append(lambda: i)
print(funcs[0](), funcs[1]())
