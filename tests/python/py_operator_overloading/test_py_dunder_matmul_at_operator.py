# vybe-test: python/py_operator_overloading/test_py_dunder_matmul_at_operator
# origin: languages/python/tests/python/test_py_operator_overloading.rs

class Matrix1D:
    def __init__(self, data):
        self.data = data

    def __matmul__(self, other):
        return sum(x * y for x, y in zip(self.data, other.data))

m1 = Matrix1D([1, 2, 3])
m2 = Matrix1D([4, 5, 6])
print(m1 @ m2)
