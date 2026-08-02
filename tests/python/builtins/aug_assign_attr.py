# vybe-test: python/builtins/aug_assign_attr
# origin: languages/python/tests/python/test_builtins.rs
# vybe-test-mode: compile

class C:
    def inc(self):
        self.count += 1
