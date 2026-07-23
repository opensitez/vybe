use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: Descriptors & Property Protocol — __get__, __set__, __delete__, __set_name__, data vs non-data descriptors
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_data_descriptor_set_get() {
    let src = r#"
class TypedInt:
    def __init__(self, name):
        self.name = name

    def __get__(self, instance, owner):
        if instance is None: return self
        return instance.__dict__.get(self.name, 0)

    def __set__(self, instance, value):
        if not isinstance(value, int):
            raise TypeError("Value must be int")
        instance.__dict__[self.name] = value

class Item:
    count = TypedInt("count")

i = Item()
i.count = 5
print(i.count)
try:
    i.count = "invalid"
except TypeError as e:
    print(e)
"#;
    assert_eq!(run_python(src), vec!["5", "Value must be int"]);
}

#[test]
fn test_py_non_data_descriptor_method_binding() {
    let src = r#"
class NonDataDesc:
    def __get__(self, instance, owner):
        if instance is None: return self
        return f"bound_to_{instance.id}"

class Host:
    def __init__(self, id):
        self.id = id
    attr = NonDataDesc()

h = Host("srv1")
print(h.attr)

# Instance attribute assignment overrides non-data descriptor!
h.__dict__["attr"] = "instance_override"
print(h.attr)
"#;
    assert_eq!(run_python(src), vec!["bound_to_srv1", "instance_override"]);
}

#[test]
fn test_py_data_descriptor_overrides_instance_dict() {
    let src = r#"
class DataDesc:
    def __get__(self, instance, owner):
        if instance is None: return self
        return "descriptor_val"

    def __set__(self, instance, value):
        pass

class Host:
    attr = DataDesc()

h = Host()
h.__dict__["attr"] = "instance_val"
# Data descriptor takes precedence over instance __dict__!
print(h.attr)
"#;
    assert_eq!(run_python(src), vec!["descriptor_val"]);
}

#[test]
fn test_py_descriptor_set_name_automatic_attr_naming() {
    let src = r#"
class AutoNamed:
    def __set_name__(self, owner, name):
        self.name = name

    def __get__(self, instance, owner):
        if instance is None: return self
        return instance.__dict__.get(self.name, None)

    def __set__(self, instance, value):
        instance.__dict__[self.name] = value

class User:
    username = AutoNamed()
    email = AutoNamed()

u = User()
u.username = "alice"
u.email = "alice@example.com"
print(u.username, u.email)
"#;
    assert_eq!(run_python(src), vec!["alice alice@example.com"]);
}

#[test]
fn test_py_property_getter_setter_deleter_full_cycle() {
    let src = r#"
class Celsius:
    def __init__(self, temp=0):
        self._temp = temp

    @property
    def temp(self):
        return self._temp

    @temp.setter
    def temp(self, value):
        if value < -273.15:
            raise ValueError("Below absolute zero")
        self._temp = value

    @temp.deleter
    def temp(self):
        print("Resetting temp to 0")
        self._temp = 0

c = Celsius(25)
print(c.temp)
c.temp = 100
print(c.temp)
del c.temp
print(c.temp)
"#;
    assert_eq!(
        run_python(src),
        vec!["25", "100", "Resetting temp to 0", "0"]
    );
}

#[test]
fn test_py_descriptor_delete_dunder_handling() {
    let src = r#"
class DeletableAttr:
    def __set_name__(self, owner, name):
        self.name = name

    def __get__(self, instance, owner):
        return instance.__dict__.get(self.name)

    def __set__(self, instance, value):
        instance.__dict__[self.name] = value

    def __delete__(self, instance):
        print(f"Deleting attribute {self.name}")
        instance.__dict__.pop(self.name, None)

class Container:
    data = DeletableAttr()

c = Container()
c.data = "payload"
print(c.data)
del c.data
print(c.data)
"#;
    assert_eq!(
        run_python(src),
        vec!["payload", "Deleting attribute data", "None"]
    );
}

#[test]
fn test_py_cached_property_functools() {
    let src = r#"
from functools import cached_property

call_count = 0

class Computation:
    @cached_property
    def expensive(self):
        global call_count
        call_count += 1
        return 42

comp = Computation()
print(comp.expensive)
print(comp.expensive)
print(call_count)  # only computed once
"#;
    assert_eq!(run_python(src), vec!["42", "42", "1"]);
}

#[test]
fn test_py_descriptor_accessed_from_class_vs_instance() {
    let src = r#"
class Desc:
    def __get__(self, instance, owner):
        if instance is None:
            return f"Desc accessed from class {owner.__name__}"
        return f"Desc accessed from instance of {owner.__name__}"

class MyClass:
    attr = Desc()

print(MyClass.attr)
print(MyClass().attr)
"#;
    assert_eq!(
        run_python(src),
        vec![
            "Desc accessed from class MyClass",
            "Desc accessed from instance of MyClass"
        ]
    );
}

#[test]
fn test_py_property_subclass_override() {
    let src = r#"
class Base:
    @property
    def val(self):
        return "base"

class Child(Base):
    @Base.val.getter
    def val(self):
        return "child"

print(Base().val)
print(Child().val)
"#;
    assert_eq!(run_python(src), vec!["base", "child"]);
}

#[test]
fn test_py_read_only_property_attribute_error() {
    let src = r#"
class ReadOnly:
    @property
    def fixed(self):
        return "immutable"

ro = ReadOnly()
print(ro.fixed)
try:
    ro.fixed = "new"
except AttributeError:
    print("AttributeError: can't set attribute")
"#;
    assert_eq!(
        run_python(src),
        vec!["immutable", "AttributeError: can't set attribute"]
    );
}
