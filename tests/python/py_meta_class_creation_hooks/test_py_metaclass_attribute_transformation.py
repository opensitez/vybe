# vybe-test: python/py_meta_class_creation_hooks/test_py_metaclass_attribute_transformation
# origin: languages/python/tests/python/test_py_meta_class_creation_hooks.rs

class UpperAttrMeta(type):
    def __new__(mcs, name, bases, attrs):
        uppercase_attrs = {}
        for key, val in attrs.items():
            if not key.startswith("__"):
                uppercase_attrs[key.upper()] = val
            else:
                uppercase_attrs[key] = val
        return super().__new__(mcs, name, bases, uppercase_attrs)

class Widget(metaclass=UpperAttrMeta):
    title = "button"
    width = 100

print(Widget.TITLE)
print(Widget.WIDTH)
print(hasattr(Widget, "title"))
