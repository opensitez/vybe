# vybe-test: python/error_handling/except_hierarchy
# origin: languages/python/tests/python/test_error_handling.rs
# vybe-test-mode: compile

try:
    pass
except ValueError:
    pass
except Exception:
    pass
except:
    pass
