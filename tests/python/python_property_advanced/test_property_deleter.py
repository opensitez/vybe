# vybe-test: python/python_property_advanced/test_property_deleter
# origin: languages/python/tests/python/test_python_property_advanced.rs

class DataHolder:
    def __init__(self, data):
        self._data = data

    @property
    def data(self):
        if self._data is None:
            raise AttributeError("data deleted")
        return self._data

    @data.deleter
    def data(self):
        self._data = None

d = DataHolder([1, 2, 3])
print(d.data)
del d.data
try:
    _ = d.data
    print("still there")
except AttributeError:
    print("deleted")
