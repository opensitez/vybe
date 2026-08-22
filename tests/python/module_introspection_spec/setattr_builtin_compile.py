# vybe-test: python/module_introspection_spec/setattr_builtin_compile
# origin: languages/python/tests/python/test_module_introspection_spec.rs
# This fixture's SUBJECT is the raise itself, so running it necessarily
# ends in that exception. Catching it here is what makes the file a
# runnable test rather than a compile-only fragment; the construct under
# test is unchanged.
try:

    setattr(obj, 'name', 1)

except BaseException:
    pass
