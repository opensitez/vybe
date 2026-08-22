# vybe-test: python/module_introspection_spec/getattr_builtin_compile
# origin: languages/python/tests/python/test_module_introspection_spec.rs
obj = 1

x = getattr(obj, 'name', None)
