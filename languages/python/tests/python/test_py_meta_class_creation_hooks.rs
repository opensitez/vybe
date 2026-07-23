use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: Metaclasses & Class Creation Hooks — type(), __new__, __prepare__, __init_subclass__, dynamic class generation
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_metaclass_attribute_transformation() {
    let src = r#"
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
"#;
    assert_eq!(run_python(src), vec!["button", "100", "False"]);
}

#[test]
fn test_py_metaclass_prepare_ordered_dict() {
    let src = r#"
class OrderedMeta(type):
    @classmethod
    def __prepare__(mcs, name, bases):
        return {"_field_order": []}

    def __new__(mcs, name, bases, attrs):
        for k in attrs:
            if not k.startswith("__"):
                attrs["_field_order"].append(k)
        return super().__new__(mcs, name, bases, attrs)

class Model(metaclass=OrderedMeta):
    id = 1
    name = "test"
    active = True

print(Model._field_order)
"#;
    assert_eq!(run_python(src), vec!["['id', 'name', 'active']"]);
}

#[test]
fn test_py_dynamic_class_creation_type_builtin() {
    let src = r#"
def init(self, name):
    self.name = name

def greet(self):
    return f"Hello {self.name}"

Person = type("Person", (object,), {
    "__init__": init,
    "greet": greet,
    "species": "Human"
})

p = Person("Alice")
print(p.greet())
print(p.species)
print(type(p).__name__)
"#;
    assert_eq!(run_python(src), vec!["Hello Alice", "Human", "Person"]);
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

s1 = Database()
s2 = Database()
print(s1 is s2)
"#;
    assert_eq!(run_python(src), vec!["True"]);
}

#[test]
fn test_py_init_subclass_attribute_validation() {
    let src = r#"
class BaseValidated:
    def __init_subclass__(cls, required_attr=None, **kwargs):
        super().__init_subclass__(**kwargs)
        if required_attr and not hasattr(cls, required_attr):
            raise TypeError(f"Class {cls.__name__} missing required attribute '{required_attr}'")

class ValidChild(BaseValidated, required_attr="version"):
    version = "1.0"

print(ValidChild.version)

try:
    class InvalidChild(BaseValidated, required_attr="missing"):
        pass
except TypeError as e:
    print(e)
"#;
    assert_eq!(
        run_python(src),
        vec![
            "1.0",
            "Class InvalidChild missing required attribute 'missing'"
        ]
    );
}

#[test]
fn test_py_metaclass_registration_pattern() {
    let src = r#"
registry = {}

class RegistryMeta(type):
    def __new__(mcs, name, bases, attrs):
        cls = super().__new__(mcs, name, bases, attrs)
        if name != "BaseHandler":
            registry[name] = cls
        return cls

class BaseHandler(metaclass=RegistryMeta): pass
class HTTPHandler(BaseHandler): pass
class FTPHandler(BaseHandler): pass

print(sorted(registry.keys()))
"#;
    assert_eq!(run_python(src), vec!["['FTPHandler', 'HTTPHandler']"]);
}

#[test]
fn test_py_metaclass_instance_check_override() {
    let src = r#"
class CustomInstanceMeta(type):
    def __instancecheck__(cls, instance):
        return hasattr(instance, "custom_protocol")

class CustomProtocolInterface(metaclass=CustomInstanceMeta): pass

class Implementation:
    custom_protocol = True

class Other: pass

print(isinstance(Implementation(), CustomProtocolInterface))
print(isinstance(Other(), CustomProtocolInterface))
"#;
    assert_eq!(run_python(src), vec!["True", "False"]);
}

#[test]
fn test_py_metaclass_subclass_check_override() {
    let src = r#"
class CustomSubclassMeta(type):
    def __subclasscheck__(cls, subclass):
        return hasattr(subclass, "is_compatible")

class Interface(metaclass=CustomSubclassMeta): pass

class CompatibleClass:
    is_compatible = True

class IncompatibleClass: pass

print(issubclass(CompatibleClass, Interface))
print(issubclass(IncompatibleClass, Interface))
"#;
    assert_eq!(run_python(src), vec!["True", "False"]);
}

#[test]
fn test_py_class_creation_namespace_modification() {
    let src = r#"
class AutoPropertyMeta(type):
    def __new__(mcs, name, bases, attrs):
        for k, v in list(attrs.items()):
            if k.startswith("get_"):
                prop_name = k[4:]
                attrs[prop_name] = property(v)
        return super().__new__(mcs, name, bases, attrs)

class User(metaclass=AutoPropertyMeta):
    def __init__(self, name):
        self._name = name
    def get_name(self):
        return self._name

u = User("Alice")
print(u.name)
"#;
    assert_eq!(run_python(src), vec!["Alice"]);
}

#[test]
fn test_py_metaclass_inheritance_consistency() {
    let src = r#"
class MetaA(type): pass
class MetaB(MetaA): pass

class ClassA(metaclass=MetaA): pass
class ClassB(ClassA, metaclass=MetaB): pass

print(type(ClassA).__name__)
print(type(ClassB).__name__)
"#;
    assert_eq!(run_python(src), vec!["MetaA", "MetaB"]);
}
