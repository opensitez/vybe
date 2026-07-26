// Python metaclass creation — type(), __init_subclass__, __class_getitem__, custom metaclass
use super::helpers::run_python;

#[test]
fn test_type_creates_class() {
    let script = r#"
MyClass = type("MyClass", (object,), {"x": 10, "greet": lambda self: "hi"})
obj = MyClass()
print(type(obj).__name__)
print(obj.x)
print(obj.greet())
"#;
    assert_eq!(run_python(script), vec!["MyClass", "10", "hi"]);
}

#[test]
fn test_metaclass_intercepts_class_creation() {
    let script = r#"
class UpperMeta(type):
    def __new__(mcs, name, bases, namespace):
        upper_ns = {k.upper(): v for k, v in namespace.items() if not k.startswith("__")}
        upper_ns.update({k: v for k, v in namespace.items() if k.startswith("__")})
        return super().__new__(mcs, name, bases, upper_ns)

class MyClass(metaclass=UpperMeta):
    value = 42

print(MyClass.VALUE)
"#;
    assert_eq!(run_python(script), vec!["42"]);
}

#[test]
fn test_init_subclass_hook() {
    let script = r#"
class Plugin:
    registry = []

    def __init_subclass__(cls, **kwargs):
        super().__init_subclass__(**kwargs)
        Plugin.registry.append(cls.__name__)

class Alpha(Plugin):
    pass

class Beta(Plugin):
    pass

print(Plugin.registry)
"#;
    assert_eq!(run_python(script), vec!["['Alpha', 'Beta']"]);
}

#[test]
fn test_metaclass_validates_subclass() {
    let script = r#"
class Singleton(type):
    _instances = {}
    def __call__(cls, *args, **kwargs):
        if cls not in cls._instances:
            cls._instances[cls] = super().__call__(*args, **kwargs)
        return cls._instances[cls]

class DB(metaclass=Singleton):
    def __init__(self):
        self.id = id(self)

a = DB()
b = DB()
print(a is b)
"#;
    assert_eq!(run_python(script), vec!["True"]);
}

#[test]
fn test_class_getitem() {
    let script = r#"
class Box:
    def __class_getitem__(cls, item):
        return f"Box[{item}]"

print(Box[int])
print(Box[str])
"#;
    assert_eq!(run_python(script), vec!["Box[<class 'int'>]", "Box[<class 'str'>]"]);
}

#[test]
fn test_type_hierarchy() {
    let script = r#"
class A:
    pass

class B(A):
    pass

class C(B):
    pass

print([t.__name__ for t in C.__mro__])
"#;
    assert_eq!(run_python(script), vec!["['C', 'B', 'A', 'object']"]);
}
