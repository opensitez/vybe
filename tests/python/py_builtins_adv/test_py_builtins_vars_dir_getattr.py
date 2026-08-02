# vybe-test: python/py_builtins_adv/test_py_builtins_vars_dir_getattr
# origin: languages/python/tests/python/test_py_builtins_adv.rs

class Person:
    def __init__(self, name, age):
        self.name = name
        self.age = age

p = Person("Alice", 30)
v = vars(p)
print(v)

d = [x for x in dir(p) if not x.startswith("_")]
print("name" in d)
print("age" in d)

print(getattr(p, "name"))
print(getattr(p, "missing", "default"))
