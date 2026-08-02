# vybe-test: python/closure_extended/closure_default_arg_fix
# origin: languages/python/tests/python/test_closure_extended.rs

funcs = []
for i in range(3):
 funcs.append(lambda i=i: i)
print(funcs[1]())
