use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: ABC + metaclasses — ABC, abstractmethod, __init_subclass__, type(), __new__, __init_subclass__, ABCMeta
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_abc_abstract_base_class_cannot_instantiate() {
    let src = r#"
from abc import ABC, abstractmethod

class Shape(ABC):
    @abstractmethod
    def area(self) -> float: ...

    @abstractmethod
    def perimeter(self) -> float: ...

try:
    Shape()
except TypeError as e:
    print("Cannot instantiate abstract class")
"#;
    assert_eq!(run_python(src), vec!["Cannot instantiate abstract class"]);
}

#[test]
fn test_py_abc_concrete_subclass_implements_all() {
    let src = r#"
from abc import ABC, abstractmethod
import math

class Shape(ABC):
    @abstractmethod
    def area(self) -> float: ...

class Circle(Shape):
    def __init__(self, r):
        self.r = r
    def area(self):
        return math.pi * self.r ** 2

c = Circle(5)
print(round(c.area(), 2))
print(isinstance(c, Shape))
"#;
    assert_eq!(run_python(src), vec!["78.54", "True"]);
}

#[test]
fn test_py_abc_partial_implementation_still_abstract() {
    let src = r#"
from abc import ABC, abstractmethod

class Base(ABC):
    @abstractmethod
    def method_a(self): ...

    @abstractmethod
    def method_b(self): ...

class PartialImpl(Base):
    def method_a(self):
        return "A"
    # method_b not implemented

try:
    PartialImpl()
except TypeError:
    print("Still abstract — method_b missing")
"#;
    assert_eq!(run_python(src), vec!["Still abstract — method_b missing"]);
}

#[test]
fn test_py_abc_abstract_property() {
    let src = r#"
from abc import ABC, abstractmethod

class Configurable(ABC):
    @property
    @abstractmethod
    def config_key(self) -> str: ...

class App(Configurable):
    @property
    def config_key(self):
        return "app.settings"

a = App()
print(a.config_key)
"#;
    assert_eq!(run_python(src), vec!["app.settings"]);
}

#[test]
fn test_py_abc_abstract_classmethod() {
    let src = r#"
from abc import ABC, abstractmethod

class Plugin(ABC):
    @classmethod
    @abstractmethod
    def get_name(cls) -> str: ...

class AudioPlugin(Plugin):
    @classmethod
    def get_name(cls) -> str:
        return "AudioPlugin"

print(AudioPlugin.get_name())
print(isinstance(AudioPlugin(), Plugin))
"#;
    assert_eq!(run_python(src), vec!["AudioPlugin", "True"]);
}

#[test]
fn test_py_abc_register_virtual_subclass() {
    let src = r#"
from abc import ABC

class Serializable(ABC):
    pass

class ThirdPartyClass:
    def serialize(self): ...

Serializable.register(ThirdPartyClass)
obj = ThirdPartyClass()
print(isinstance(obj, Serializable))
print(issubclass(ThirdPartyClass, Serializable))
"#;
    assert_eq!(run_python(src), vec!["True", "True"]);
}

#[test]
fn test_py_metaclass_type_creates_class_dynamically() {
    let src = r#"
DynamicClass = type("DynamicClass", (object,), {
    "greeting": "Hello",
    "greet": lambda self: f"{self.greeting}, World!"
})

obj = DynamicClass()
print(obj.greet())
print(type(obj).__name__)
"#;
    assert_eq!(run_python(src), vec!["Hello, World!", "DynamicClass"]);
}

#[test]
fn test_py_metaclass_custom_metaclass() {
    let src = r#"
class UpperAttrMeta(type):
    def __new__(mcs, name, bases, namespace):
        upper_attrs = {
            k.upper() if not k.startswith('_') else k: v
            for k, v in namespace.items()
        }
        return super().__new__(mcs, name, bases, upper_attrs)

class MyClass(metaclass=UpperAttrMeta):
    greeting = "hello"
    count = 42

print(hasattr(MyClass, 'GREETING'))
print(MyClass.GREETING)
print(MyClass.COUNT)
"#;
    assert_eq!(run_python(src), vec!["True", "hello", "42"]);
}

#[test]
fn test_py_metaclass_singleton_pattern() {
    let src = r#"
class SingletonMeta(type):
    _instances = {}

    def __call__(cls, *args, **kwargs):
        if cls not in cls._instances:
            cls._instances[cls] = super().__call__(*args, **kwargs)
        return cls._instances[cls]

class Database(metaclass=SingletonMeta):
    def __init__(self):
        self.connected = True

db1 = Database()
db2 = Database()
print(db1 is db2)
print(db1.connected)
"#;
    assert_eq!(run_python(src), vec!["True", "True"]);
}

#[test]
fn test_py_init_subclass_hook() {
    let src = r#"
class Base:
    subclasses = []

    def __init_subclass__(cls, required_field=None, **kwargs):
        super().__init_subclass__(**kwargs)
        if required_field is None:
            raise TypeError(f"{cls.__name__} must specify required_field")
        Base.subclasses.append(cls.__name__)

class ChildA(Base, required_field="x"):
    pass

class ChildB(Base, required_field="y"):
    pass

print(Base.subclasses)
"#;
    assert_eq!(run_python(src), vec!["['ChildA', 'ChildB']"]);
}

#[test]
fn test_py_metaclass_prepare_ordered_namespace() {
    let src = r#"
from collections import OrderedDict

class OrderedMeta(type):
    @classmethod
    def __prepare__(mcs, name, bases):
        return OrderedDict()

    def __new__(mcs, name, bases, namespace):
        cls = super().__new__(mcs, name, bases, dict(namespace))
        cls._field_order = list(namespace.keys())
        return cls

class Model(metaclass=OrderedMeta):
    first_name = None
    last_name = None
    age = None

print(Model._field_order)
"#;
    assert_eq!(run_python(src), vec!["['first_name', 'last_name', 'age']"]);
}

#[test]
fn test_py_abc_abstractstaticmethod() {
    let src = r#"
from abc import ABC, abstractmethod

class Parser(ABC):
    @staticmethod
    @abstractmethod
    def parse(text: str) -> dict: ...

class JSONParser(Parser):
    @staticmethod
    def parse(text: str) -> dict:
        import json
        return json.loads(text)

result = JSONParser.parse('{"key": "value"}')
print(result["key"])
"#;
    assert_eq!(run_python(src), vec!["value"]);
}

#[test]
fn test_py_metaclass_class_creation_hooks() {
    let src = r#"
log = []

class LoggingMeta(type):
    def __new__(mcs, name, bases, namespace):
        log.append(f"Creating {name}")
        return super().__new__(mcs, name, bases, namespace)

class A(metaclass=LoggingMeta):
    pass

class B(A):
    pass

print(log)
"#;
    assert_eq!(run_python(src), vec!["['Creating A', 'Creating B']"]);
}
