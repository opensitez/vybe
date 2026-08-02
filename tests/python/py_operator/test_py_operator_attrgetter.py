# vybe-test: python/py_operator/test_py_operator_attrgetter
# origin: languages/python/tests/python/test_py_operator.rs

from operator import attrgetter

class Person:
    def __init__(self, name, age):
        self.name = name
        self.age = age

people = [Person("Bob", 25), Person("Alice", 30), Person("Charlie", 20)]
get_name = attrgetter("name")
print(get_name(people[0]))

sorted_by_age = sorted(people, key=attrgetter("age"))
print([p.name for p in sorted_by_age])
