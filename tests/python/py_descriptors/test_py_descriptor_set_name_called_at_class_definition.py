# vybe-test: python/py_descriptors/test_py_descriptor_set_name_called_at_class_definition
# origin: languages/python/tests/python/test_py_descriptors.rs

class Named:
    def __set_name__(self, owner, name):
        self.attr_name = name
        print(f"Registered: {name} on {owner.__name__}")

    def __get__(self, obj, objtype=None):
        return self.attr_name

class Widget:
    color = Named()
    size = Named()

w = Widget()
print(w.color)
print(w.size)
