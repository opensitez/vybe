# vybe-test: python/closure_extended/closure_late_binding
# origin: languages/python/tests/python/test_closure_extended.rs

funcs = []
for i in range(3):
 funcs.append(lambda: i)
print(funcs[2]())
