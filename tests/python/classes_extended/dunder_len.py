# vybe-test: python/classes_extended/dunder_len
# origin: languages/python/tests/python/test_classes_extended.rs

class Bag:
    def __init__(self):
        self.items = []
    def __len__(self):
        return len(self.items)
