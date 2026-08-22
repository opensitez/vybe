# vybe-test: python/builtins/aug_assign_attr
# origin: languages/python/tests/python/test_builtins.rs

class C:
    def inc(self):
        self.count += 1
