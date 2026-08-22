# vybe-test: python/module_introspection_spec/vars_builtin_compile
# origin: languages/python/tests/python/test_module_introspection_spec.rs

class C:
    pass
c = C()
v = vars(c)
