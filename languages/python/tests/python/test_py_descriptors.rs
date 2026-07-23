use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: Descriptors — __get__, __set__, __delete__, data vs non-data, __set_name__
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_descriptor_non_data_get_only() {
    let src = r#"
class LazyDescriptor:
    def __get__(self, obj, objtype=None):
        if obj is None:
            return self
        return 42  # computed value

class MyClass:
    value = LazyDescriptor()

m = MyClass()
print(m.value)
print(MyClass.value)  # returns the descriptor itself when accessed on class
"#;
    assert_eq!(
        run_python(src),
        vec!["42", "<__main__.LazyDescriptor object at ..."]
    );
}

#[test]
fn test_py_descriptor_data_descriptor_set_and_get() {
    let src = r#"
class PositiveInt:
    def __set_name__(self, owner, name):
        self.name = name
        self.private_name = f"_{name}"

    def __get__(self, obj, objtype=None):
        if obj is None:
            return self
        return getattr(obj, self.private_name, 0)

    def __set__(self, obj, value):
        if value < 0:
            raise ValueError(f"{self.name} must be positive")
        setattr(obj, self.private_name, value)

class Product:
    quantity = PositiveInt()

    def __init__(self, qty):
        self.quantity = qty

p = Product(10)
print(p.quantity)
try:
    p.quantity = -5
except ValueError as e:
    print(e)
"#;
    assert_eq!(run_python(src), vec!["10", "quantity must be positive"]);
}

#[test]
fn test_py_descriptor_set_name_called_at_class_definition() {
    let src = r#"
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
"#;
    assert_eq!(
        run_python(src),
        vec![
            "Registered: color on Widget",
            "Registered: size on Widget",
            "color",
            "size"
        ]
    );
}

#[test]
fn test_py_descriptor_delete_protocol() {
    let src = r#"
class Tracked:
    def __init__(self):
        self._data = {}

    def __set_name__(self, owner, name):
        self.name = name

    def __get__(self, obj, objtype=None):
        if obj is None:
            return self
        return obj.__dict__.get(f"_val_{self.name}")

    def __set__(self, obj, value):
        obj.__dict__[f"_val_{self.name}"] = value

    def __delete__(self, obj):
        print(f"Deleting {self.name}")
        obj.__dict__.pop(f"_val_{self.name}", None)

class User:
    name = Tracked()

u = User()
u.name = "Alice"
print(u.name)
del u.name
print(u.name)
"#;
    assert_eq!(run_python(src), vec!["Alice", "Deleting name", "None"]);
}

#[test]
fn test_py_descriptor_property_is_data_descriptor() {
    let src = r#"
class Celsius:
    def __init__(self):
        self._temp = 0.0

    @property
    def temperature(self):
        return self._temp

    @temperature.setter
    def temperature(self, value):
        if value < -273.15:
            raise ValueError("Temperature below absolute zero!")
        self._temp = value

c = Celsius()
c.temperature = 100
print(c.temperature)
try:
    c.temperature = -300
except ValueError as e:
    print(e)
"#;
    assert_eq!(
        run_python(src),
        vec!["100", "Temperature below absolute zero!"]
    );
}

#[test]
fn test_py_descriptor_shadowing_by_instance_dict_non_data() {
    let src = r#"
class NonDataDesc:
    def __get__(self, obj, objtype=None):
        return "from_descriptor"

class MyClass:
    value = NonDataDesc()

m = MyClass()
print(m.value)
# Instance dict can shadow non-data descriptor
m.__dict__['value'] = "from_instance"
print(m.value)  # instance dict wins over non-data descriptor
"#;
    assert_eq!(run_python(src), vec!["from_descriptor", "from_instance"]);
}

#[test]
fn test_py_descriptor_data_descriptor_takes_priority_over_instance_dict() {
    let src = r#"
class DataDesc:
    def __get__(self, obj, objtype=None):
        if obj is None:
            return self
        return "from_data_descriptor"

    def __set__(self, obj, value):
        print(f"set called with {value}")
        obj.__dict__['_hidden'] = value

class MyClass:
    attr = DataDesc()

m = MyClass()
m.attr = "test"
m.__dict__['attr'] = "direct_dict"  # try to shadow
print(m.attr)  # data descriptor still wins!
"#;
    assert_eq!(
        run_python(src),
        vec!["set called with test", "from_data_descriptor"]
    );
}

#[test]
fn test_py_descriptor_classmethod_staticmethod_are_descriptors() {
    let src = r#"
class Foo:
    @classmethod
    def cls_method(cls):
        return cls.__name__

    @staticmethod
    def static_method(x):
        return x * 2

# classmethod and staticmethod are implemented as descriptors
print(isinstance(Foo.__dict__['cls_method'], classmethod))
print(isinstance(Foo.__dict__['static_method'], staticmethod))
print(Foo.cls_method())
print(Foo.static_method(5))
"#;
    assert_eq!(run_python(src), vec!["True", "True", "Foo", "10"]);
}

#[test]
fn test_py_descriptor_functions_are_non_data_descriptors() {
    let src = r#"
class MyClass:
    def greet(self):
        return f"Hello from {type(self).__name__}"

m = MyClass()
# Functions are non-data descriptors — accessing via instance binds self
unbound = MyClass.__dict__['greet']
bound = m.greet
print(type(unbound).__name__)
print(type(bound).__name__)
print(bound())
"#;
    assert_eq!(
        run_python(src),
        vec!["function", "method", "Hello from MyClass"]
    );
}

#[test]
fn test_py_descriptor_shared_storage_for_multiple_instances() {
    let src = r#"
class SlotDescriptor:
    def __set_name__(self, owner, name):
        self.slot_key = f"_slot_{name}"

    def __get__(self, obj, objtype=None):
        if obj is None:
            return self
        return getattr(obj, self.slot_key, "default")

    def __set__(self, obj, value):
        setattr(obj, self.slot_key, value)

class Widget:
    color = SlotDescriptor()
    size = SlotDescriptor()

w1, w2 = Widget(), Widget()
w1.color = "red"
w2.color = "blue"
w1.size = "small"
print(w1.color, w2.color, w1.size, w2.size)
"#;
    assert_eq!(run_python(src), vec!["red blue small default"]);
}

#[test]
fn test_py_descriptor_caching_computed_value() {
    let src = r#"
class CachedProperty:
    def __init__(self, func):
        self.func = func
        self.name = None

    def __set_name__(self, owner, name):
        self.name = name

    def __get__(self, obj, objtype=None):
        if obj is None:
            return self
        if self.name not in obj.__dict__:
            obj.__dict__[self.name] = self.func(obj)
        return obj.__dict__[self.name]

call_count = [0]

class DataModel:
    @CachedProperty
    def processed(self):
        call_count[0] += 1
        return 42 * 2

d = DataModel()
print(d.processed)
print(d.processed)
print(d.processed)
print(call_count[0])  # only computed once!
"#;
    assert_eq!(run_python(src), vec!["84", "84", "84", "1"]);
}

#[test]
fn test_py_descriptor_inheritance_override() {
    let src = r#"
class ValidatedField:
    def __set_name__(self, owner, name):
        self.name = name

    def __get__(self, obj, objtype=None):
        if obj is None:
            return self
        return obj.__dict__.get(f"_{self.name}")

    def __set__(self, obj, value):
        value = self.validate(value)
        obj.__dict__[f"_{self.name}"] = value

    def validate(self, value):
        return value

class UpperField(ValidatedField):
    def validate(self, value):
        return value.upper()

class Model:
    name = UpperField()

m = Model()
m.name = "hello"
print(m.name)
"#;
    assert_eq!(run_python(src), vec!["HELLO"]);
}

#[test]
fn test_py_descriptor_with_weakref_storage() {
    let src = r#"
import weakref

class WeakrefDescriptor:
    def __set_name__(self, owner, name):
        self.name = name
        self._refs = weakref.WeakValueDictionary()

    def __get__(self, obj, objtype=None):
        if obj is None:
            return self
        return self._refs.get(id(obj))

    def __set__(self, obj, value):
        self._refs[id(obj)] = value

class Container:
    data = WeakrefDescriptor()

class Payload:
    pass

c = Container()
p = Payload()
c.data = p
print(c.data is p)
"#;
    assert_eq!(run_python(src), vec!["True"]);
}
