# vybe-test: python/closure_extended/closure_cell_independent
# origin: languages/python/tests/python/test_closure_extended.rs

def make():
 return lambda: 1
a = make()
b = make()
print(a(), b())
