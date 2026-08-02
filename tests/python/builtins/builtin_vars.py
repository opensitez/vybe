# vybe-test: python/builtins/builtin_vars
# origin: languages/python/tests/python/test_builtins.rs
# vybe-test-mode: compile

class C:
    def __init__(self):
        self.x = 1
c = C()
d = vars(c)
