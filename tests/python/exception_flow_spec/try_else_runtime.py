# vybe-test: python/exception_flow_spec/try_else_runtime
# origin: languages/python/tests/python/test_exception_flow_spec.rs

try:
    x = 1
except Exception:
    print('except')
else:
    print('else')
