# vybe-test: python/python_lambda_closures/test_closure_late_binding_fixed
# origin: languages/python/tests/python/test_python_lambda_closures.rs

# classic loop closure bug — fix using default arg
funcs = [lambda x, i=i: x + i for i in range(3)]
print(funcs[0](10))
print(funcs[1](10))
print(funcs[2](10))
