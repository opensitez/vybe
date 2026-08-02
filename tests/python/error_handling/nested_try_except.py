# vybe-test: python/error_handling/nested_try_except
# origin: languages/python/tests/python/test_error_handling.rs
# vybe-test-mode: compile

try:
    try:
        risky()
    except ValueError:
        pass
except Exception:
    pass
