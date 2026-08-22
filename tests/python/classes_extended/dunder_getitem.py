# vybe-test: python/classes_extended/dunder_getitem
# origin: languages/python/tests/python/test_classes_extended.rs

class Row:
    def __init__(self, data):
        self.data = data
    def __getitem__(self, key):
        return self.data[key]
