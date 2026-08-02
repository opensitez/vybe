# vybe-test: python/py_decorators/test_py_closure_late_binding_fix_with_default_arg
# origin: languages/python/tests/python/test_py_decorators.rs

# Without default arg, late binding captures final loop variable
funcs_late = [lambda: i for i in range(3)]
print([f() for f in funcs_late])

# Fixed with default parameter default arg
funcs_fixed = [lambda i=i: i for i in range(3)]
print([f() for f in funcs_fixed])
