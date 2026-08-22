# vybe-test: python/classes/user_method_named_append
# origin: languages/python/tests/python/test_classes.rs

class C:
    def append(self, x):
        print(x)
c = C()
c.append(42)
