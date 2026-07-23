use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: Meta Dunders & Object Customization — __str__, __repr__, __eq__, __hash__, __dir__, __sizeof__, __bytes__, __init_subclass__
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_dunder_str_vs_repr_formatting() {
    let src = r#"
class Person:
    def __init__(self, name):
        self.name = name

    def __str__(self):
        return f"Person({self.name})"

    def __repr__(self):
        return f"<Person name={self.name!r}>"

p = Person("Alice")
print(str(p))
print(repr(p))
print(f"{p}")
print(f"{p!r}")
"#;
    assert_eq!(
        run_python(src),
        vec![
            "Person(Alice)",
            "<Person name='Alice'>",
            "Person(Alice)",
            "<Person name='Alice'>"
        ]
    );
}

#[test]
fn test_py_dunder_eq_and_hash_consistency() {
    let src = r#"
class User:
    def __init__(self, user_id, name):
        self.user_id = user_id
        self.name = name

    def __eq__(self, other):
        if not isinstance(other, User): return False
        return self.user_id == other.user_id

    def __hash__(self):
        return hash(self.user_id)

u1 = User(1, "Alice")
u2 = User(1, "Alice Modified")
print(u1 == u2)
print(hash(u1) == hash(u2))

s = {u1, u2}
print(len(s))
"#;
    assert_eq!(run_python(src), vec!["True", "True", "1"]);
}

#[test]
fn test_py_unhashable_type_on_mutable_eq_override() {
    let src = r#"
class MutableContainer:
    def __init__(self, val):
        self.val = val

    def __eq__(self, other):
        return self.val == other.val

c = MutableContainer(10)
try:
    hash(c)
except TypeError:
    print("TypeError: unhashable type")
"#;
    assert_eq!(run_python(src), vec!["TypeError: unhashable type"]);
}

#[test]
fn test_py_dunder_dir_custom_attribute_listing() {
    let src = r#"
class CustomDir:
    def __dir__(self):
        return ["custom_attr_a", "custom_attr_b"]

c = CustomDir()
print(dir(c))
"#;
    assert_eq!(run_python(src), vec!["['custom_attr_a', 'custom_attr_b']"]);
}

#[test]
fn test_py_dunder_bytes_conversion() {
    let src = r#"
class Serializable:
    def __bytes__(self):
        return b"custom_bytes_repr"

s = Serializable()
print(bytes(s))
"#;
    assert_eq!(run_python(src), vec!["b'custom_bytes_repr'"]);
}

#[test]
fn test_py_dunder_sizeof_override() {
    let src = r#"
import sys

class HugeVirtualObject:
    def __sizeof__(self):
        return 1024 * 1024  # 1MB

obj = HugeVirtualObject()
print(sys.getsizeof(obj))
"#;
    assert_eq!(run_python(src), vec!["1048576"]);
}

#[test]
fn test_py_init_subclass_automatic_registration() {
    let src = r#"
class PluginRegistry:
    plugins = {}

    def __init_subclass__(cls, plugin_name=None, **kwargs):
        super().__init_subclass__(**kwargs)
        if plugin_name:
            cls.plugins[plugin_name] = cls

class AudioPlugin(PluginRegistry, plugin_name="audio"):
    pass

class VideoPlugin(PluginRegistry, plugin_name="video"):
    pass

print(sorted(PluginRegistry.plugins.keys()))
"#;
    assert_eq!(run_python(src), vec!["['audio', 'video']"]);
}

#[test]
fn test_py_dunder_format_custom_specifier() {
    let src = r#"
class HexInt:
    def __init__(self, val):
        self.val = val

    def __format__(self, format_spec):
        if format_spec == "x":
            return hex(self.val)
        return str(self.val)

h = HexInt(255)
print(f"{h:x}")
print(f"{h}")
"#;
    assert_eq!(run_python(src), vec!["0xff", "255"]);
}

#[test]
fn test_py_dunder_iter_and_next_protocol() {
    let src = r#"
class Countdown:
    def __init__(self, start):
        self.current = start

    def __iter__(self):
        return self

    def __next__(self):
        if self.current <= 0:
            raise StopIteration
        val = self.current
        self.current -= 1
        return val

print(list(Countdown(3)))
"#;
    assert_eq!(run_python(src), vec!["[3, 2, 1]"]);
}

#[test]
fn test_py_dunder_del_destructor() {
    let src = r#"
log = []

class AutoClean:
    def __del__(self):
        log.append("cleaned")

obj = AutoClean()
del obj
import gc; gc.collect()
print(log)
"#;
    assert_eq!(run_python(src), vec!["['cleaned']"]);
}
