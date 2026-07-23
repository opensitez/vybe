use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: Class Attributes & Descriptors — slots, property, classmethod, staticmethod, getattr, setattr, hasattr, delattr
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_getattr_custom_fallback() {
    let src = r#"
class DynamicAttributes:
    def __init__(self):
        self.known = "exists"

    def __getattr__(self, item):
        return f"dynamic_{item}"

d = DynamicAttributes()
print(d.known)
print(d.unknown_property)
"#;
    assert_eq!(run_python(src), vec!["exists", "dynamic_unknown_property"]);
}

#[test]
fn test_py_getattribute_interception() {
    let src = r#"
class Interceptor:
    def __init__(self):
        self.val = 42

    def __getattribute__(self, item):
        if item == "secret":
            return "intercepted"
        return object.__getattribute__(self, item)

i = Interceptor()
print(i.val)
print(i.secret)
"#;
    assert_eq!(run_python(src), vec!["42", "intercepted"]);
}

#[test]
fn test_py_setattr_delattr_custom_validation() {
    let src = r#"
class Restricted:
    def __setattr__(self, key, value):
        if key.startswith("_"):
            raise AttributeError("Private key not allowed")
        self.__dict__[key] = value

    def __delattr__(self, key):
        print(f"Deleting {key}")
        object.__delattr__(self, key)

r = Restricted()
r.name = "Public"
print(r.name)
del r.name

try:
    r._secret = 123
except AttributeError as e:
    print(e)
"#;
    assert_eq!(
        run_python(src),
        vec!["Public", "Deleting name", "Private key not allowed"]
    );
}

#[test]
fn test_py_property_cached_pattern() {
    let src = r#"
class Circle:
    def __init__(self, radius):
        self.radius = radius

    @property
    def area(self):
        return 3.14159 * (self.radius ** 2)

c = Circle(10)
print(round(c.area, 2))
c.radius = 5
print(round(c.area, 2))
"#;
    assert_eq!(run_python(src), vec!["314.16", "78.54"]);
}

#[test]
fn test_py_classmethod_factory_constructor() {
    let src = r#"
class Person:
    def __init__(self, name, age):
        self.name = name
        self.age = age

    @classmethod
    def from_birth_year(cls, name, birth_year, current_year=2024):
        return cls(name, current_year - birth_year)

p = Person.from_birth_year("Alice", 1994)
print(p.name, p.age)
"#;
    assert_eq!(run_python(src), vec!["Alice 30"]);
}

#[test]
fn test_py_staticmethod_pure_utility() {
    let src = r#"
class MathHelper:
    @staticmethod
    def is_even(n):
        return n % 2 == 0

print(MathHelper.is_even(4))
print(MathHelper.is_even(7))
"#;
    assert_eq!(run_python(src), vec!["True", "False"]);
}

#[test]
fn test_py_slots_memory_optimization() {
    let src = r#"
class PointSlotted:
    __slots__ = ("x", "y")

    def __init__(self, x, y):
        self.x = x
        self.y = y

p = PointSlotted(1, 2)
print(p.x, p.y)
print(hasattr(p, "__dict__"))

try:
    p.z = 3
except AttributeError:
    print("AttributeError: no z in slots")
"#;
    assert_eq!(
        run_python(src),
        vec!["1 2", "False", "AttributeError: no z in slots"]
    );
}

#[test]
fn test_py_hasattr_getattr_setattr_builtins() {
    let src = r#"
class Dummy: pass

d = Dummy()
setattr(d, "dynamic_key", "value_123")
print(hasattr(d, "dynamic_key"))
print(getattr(d, "dynamic_key"))
delattr(d, "dynamic_key")
print(hasattr(d, "dynamic_key"))
"#;
    assert_eq!(run_python(src), vec!["True", "value_123", "False"]);
}

#[test]
fn test_py_descriptor_protocol_set_name() {
    let src = r#"
class Val:
    def __set_name__(self, owner, name):
        self.name = name

    def __get__(self, instance, owner):
        if instance is None: return self
        return instance.__dict__.get(self.name, 0)

    def __set__(self, instance, value):
        instance.__dict__[self.name] = value

class Widget:
    width = Val()
    height = Val()

w = Widget()
w.width = 100
w.height = 200
print(w.width, w.height)
"#;
    assert_eq!(run_python(src), vec!["100 200"]);
}

#[test]
fn test_py_class_dict_vs_instance_dict() {
    let src = r#"
class Demo:
    class_attr = 10

d = Demo()
print(d.class_attr)
print("class_attr" in d.__dict__)
print("class_attr" in Demo.__dict__)
"#;
    assert_eq!(run_python(src), vec!["10", "False", "True"]);
}
