# vybe-test: python/py_generators_iterators/test_py_iterable_vs_iterator_distinction
# origin: languages/python/tests/python/test_py_generators_iterators.rs

class NumberList:
    def __init__(self, data):
        self.data = data

    def __iter__(self):
        return iter(self.data)

nl = NumberList([10, 20, 30])
for v in nl:
    print(v)
# Can iterate multiple times:
print(sum(nl))
