# vybe-test: python/exception_flow_spec/nested_except_runtime
# origin: languages/python/tests/python/test_exception_flow_spec.rs

try:
    try:
        raise ValueError()
    except ValueError:
        print('inner')
except Exception:
    print('outer')
