# vybe-test: python/programs/counter_class
# origin: languages/python/tests/python/test_programs.rs
# vybe-test-mode: compile

class Counter:
    def __init__(self):
        self.count = 0

    def increment(self):
        self.count += 1

    def get(self):
        return self.count
