# vybe-test: python/python_sorted_key_reverse/test_sorted_custom_objects
# origin: languages/python/tests/python/test_python_sorted_key_reverse.rs

class Person:
    def __init__(self, name, age):
        self.name = name
        self.age = age
    def __repr__(self):
        return self.name

people = [Person("Bob", 30), Person("Alice", 25), Person("Carol", 35)]
by_age = sorted(people, key=lambda p: p.age)
print([p.name for p in by_age])
