use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: Descriptor & Property Chaining — property getter/setter/deleter, cached_property, descriptor precedence, __set_name__
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_property_full_descriptor_lifecycle() {
    let src = r#"
class Temperature:
    def __init__(self, celsius=0.0):
        self._celsius = celsius

    @property
    def celsius(self):
        return self._celsius

    @celsius.setter
    def celsius(self, value):
        if value < -273.15:
            raise ValueError("Temperature below absolute zero")
        self._celsius = value

    @celsius.deleter
    def celsius(self):
        print("Resetting temperature")
        self._celsius = 0.0

t = Temperature(25.0)
print(t.celsius)
t.celsius = 100.0
print(t.celsius)
del t.celsius
print(t.celsius)
"#;
    assert_eq!(
        run_python(src),
        vec!["25.0", "100.0", "Resetting temperature", "0.0"]
    );
}

#[test]
fn test_py_data_descriptor_takes_precedence_over_instance_dict() {
    let src = r#"
class OverrideDesc:
    def __get__(self, instance, owner):
        if instance is None: return self
        return "descriptor_value"

    def __set__(self, instance, value):
        instance.__dict__["override"] = "intercepted"

class Host:
    override = OverrideDesc()

h = Host()
h.__dict__["override"] = "direct_dict_value"
# Data descriptor __get__ takes precedence over instance __dict__!
print(h.override)
"#;
    assert_eq!(run_python(src), vec!["descriptor_value"]);
}

#[test]
fn test_py_non_data_descriptor_yields_to_instance_dict() {
    let src = r#"
class MethodLikeDesc:
    def __get__(self, instance, owner):
        if instance is None: return self
        return "non_data_val"

class Host:
    attr = MethodLikeDesc()

h = Host()
print(h.attr)
# Instance assignment overrides non-data descriptor!
h.__dict__["attr"] = "instance_val"
print(h.attr)
"#;
    assert_eq!(run_python(src), vec!["non_data_val", "instance_val"]);
}

#[test]
fn test_py_descriptor_set_name_automatic_binding() {
    let src = r#"
class Field:
    def __set_name__(self, owner, name):
        self.name = f"_{name}"

    def __get__(self, instance, owner):
        if instance is None: return self
        return getattr(instance, self.name, None)

    def __set__(self, instance, value):
        setattr(instance, self.name, value)

class Person:
    first_name = Field()
    last_name = Field()

p = Person()
p.first_name = "John"
p.last_name = "Doe"
print(p.first_name, p.last_name)
print(p._first_name)
"#;
    assert_eq!(run_python(src), vec!["John Doe", "John"]);
}

#[test]
fn test_py_functools_cached_property_memoization() {
    let src = r#"
from functools import cached_property

class HeavyComputation:
    def __init__(self, data):
        self.data = data

    @cached_property
    def total(self):
        print("Computing total")
        return sum(self.data)

hc = HeavyComputation([1, 2, 3, 4, 5])
print(hc.total)
print(hc.total)  # cached in instance __dict__!
print("total" in hc.__dict__)
"#;
    assert_eq!(run_python(src), vec!["Computing total", "15", "15", "True"]);
}

#[test]
fn test_py_property_read_only_attribute_error() {
    let src = r#"
class ReadOnly:
    @property
    def secret(self):
        return 42

ro = ReadOnly()
print(ro.secret)
try:
    ro.secret = 100
except AttributeError:
    print("AttributeError: can't set attribute")
"#;
    assert_eq!(
        run_python(src),
        vec!["42", "AttributeError: can't set attribute"]
    );
}

#[test]
fn test_py_descriptor_delete_hook() {
    let src = r#"
class ManagedAttr:
    def __set_name__(self, owner, name):
        self.name = name

    def __get__(self, instance, owner):
        return instance.__dict__.get(self.name)

    def __set__(self, instance, value):
        instance.__dict__[self.name] = value

    def __delete__(self, instance):
        print(f"Deleted {self.name}")
        instance.__dict__.pop(self.name, None)

class Container:
    item = ManagedAttr()

c = Container()
c.item = "hello"
print(c.item)
del c.item
print(c.item)
"#;
    assert_eq!(run_python(src), vec!["hello", "Deleted item", "None"]);
}

#[test]
fn test_py_property_subclass_getter_override() {
    let src = r#"
class Base:
    @property
    def label(self):
        return "base_label"

class Child(Base):
    @Base.label.getter
    def label(self):
        return "child_label"

print(Base().label)
print(Child().label)
"#;
    assert_eq!(run_python(src), vec!["base_label", "child_label"]);
}

#[test]
fn test_py_class_method_descriptor_binding() {
    let src = r#"
class CustomMethod:
    def __init__(self, func):
        self.func = func

    def __get__(self, instance, owner):
        if instance is None: return self
        return lambda *args: self.func(instance, *args)

class Calculator:
    def __init__(self, base):
        self.base = base

    @CustomMethod
    def add(self, x):
        return self.base + x

calc = Calculator(10)
print(calc.add(5))
"#;
    assert_eq!(run_python(src), vec!["15"]);
}

#[test]
fn test_py_descriptor_class_access_returns_descriptor_itself() {
    let src = r#"
class SimpleDesc:
    def __get__(self, instance, owner):
        if instance is None:
            return "class_access"
        return "instance_access"

class Host:
    d = SimpleDesc()

print(Host.d)
print(Host().d)
"#;
    assert_eq!(run_python(src), vec!["class_access", "instance_access"]);
}
