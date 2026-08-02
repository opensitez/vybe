# vybe-test: python/closure_truthiness_spec/nonlocal_two_levels_compile
# origin: languages/python/tests/python/test_closure_truthiness_spec.rs
# vybe-test-mode: compile

def outer():
    x = 0
    def mid():
        def inner():
            nonlocal x
            x += 1
        inner()
