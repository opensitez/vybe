# vybe-test: python/python_operator_module/test_operator_attrgetter
# origin: languages/python/tests/python/test_python_operator_module.rs

import operator

class Person:
    def __init__(self, name, age):
        self.name = name
        self.age = age

people = [Person('Bob', 30), Person('Alice', 25), Person('Carol', 35)]
by_name = sorted(people, key=operator.attrgetter('name'))
print([p.name for p in by_name])
