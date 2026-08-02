# vybe-test: python/classes/user_method_named_get
# origin: languages/python/tests/python/test_classes.rs
# vybe-test-mode: compile

class C:
    def get(self):
        return 42
c = C()
print(c.get())
