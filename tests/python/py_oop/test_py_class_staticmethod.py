# vybe-test: python/py_oop/test_py_class_staticmethod
# origin: languages/python/tests/python/test_py_oop.rs

class MathUtils:
    @staticmethod
    def gcd(a, b):
        while b:
            a, b = b, a % b
        return a

    @staticmethod
    def lcm(a, b):
        return a * b // MathUtils.gcd(a, b)

print(MathUtils.gcd(12, 8))
print(MathUtils.lcm(4, 6))
obj = MathUtils()
print(obj.gcd(15, 5))  # also callable via instance
