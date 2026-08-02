# vybe-test: python/global_nonlocal/closure_capture_loop_with_default_fix
# origin: languages/python/tests/python/test_global_nonlocal.rs

funcs = [lambda x=i: x for i in range(2)]
print(funcs[0](), funcs[1]())
