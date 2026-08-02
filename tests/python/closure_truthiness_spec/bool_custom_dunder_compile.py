# vybe-test: python/closure_truthiness_spec/bool_custom_dunder_compile
# origin: languages/python/tests/python/test_closure_truthiness_spec.rs
# vybe-test-mode: compile

class Flag:
    def __bool__(self):
        return False
