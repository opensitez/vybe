use super::helpers::run_python;

// types — SimpleNamespace, MappingProxyType, new_class, resolve_bases, MethodType, FunctionType, DynamicClassAttribute

#[test]
fn test_types_simple_namespace_attribute_access() {
    let out = run_python(
        r#"
from types import SimpleNamespace
ns = SimpleNamespace(name="Alice", age=30, city="Paris")
print(ns.name)
print(ns.age)
print(ns.city)
"#,
    );
    assert_eq!(out, vec!["Alice", "30", "Paris"]);
}

#[test]
fn test_types_simple_namespace_mutation_and_addition() {
    let out = run_python(
        r#"
from types import SimpleNamespace
ns = SimpleNamespace(x=10)
ns.x = 20
ns.y = 30
del ns.x
print(hasattr(ns, "x"))
print(ns.y)
"#,
    );
    assert_eq!(out, vec!["False", "30"]);
}

#[test]
fn test_types_simple_namespace_repr() {
    let out = run_python(
        r#"
from types import SimpleNamespace
ns = SimpleNamespace(a=1, b="test")
print(repr(ns))
"#,
    );
    assert_eq!(out, vec!["namespace(a=1, b='test')"]);
}

#[test]
fn test_types_simple_namespace_equality() {
    let out = run_python(
        r#"
from types import SimpleNamespace
ns1 = SimpleNamespace(a=1, b=2)
ns2 = SimpleNamespace(a=1, b=2)
ns3 = SimpleNamespace(a=1, b=3)
print(ns1 == ns2)
print(ns1 == ns3)
"#,
    );
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_types_mapping_proxy_type_read_only() {
    let out = run_python(
        r#"
from types import MappingProxyType
d = {"key": "val"}
proxy = MappingProxyType(d)
print(proxy["key"])
print("key" in proxy)
print(len(proxy))
"#,
    );
    assert_eq!(out, vec!["val", "True", "1"]);
}

#[test]
fn test_types_mapping_proxy_type_mutation_fails() {
    let out = run_python(
        r#"
from types import MappingProxyType
d = {"a": 1}
proxy = MappingProxyType(d)
try:
    proxy["a"] = 2
except TypeError:
    print("TypeError")
"#,
    );
    assert_eq!(out, vec!["TypeError"]);
}

#[test]
fn test_types_mapping_proxy_type_underlying_dict_mutation_reflected() {
    let out = run_python(
        r#"
from types import MappingProxyType
d = {"count": 10}
proxy = MappingProxyType(d)
print(proxy["count"])
d["count"] = 20
print(proxy["count"])
"#,
    );
    assert_eq!(out, vec!["10", "20"]);
}

#[test]
fn test_types_mapping_proxy_type_keys_values_items() {
    let out = run_python(
        r#"
from types import MappingProxyType
d = {"a": 1, "b": 2}
proxy = MappingProxyType(d)
print(list(proxy.keys()))
print(list(proxy.values()))
print(list(proxy.items()))
"#,
    );
    assert_eq!(out, vec!["['a', 'b']", "[1, 2]", "[('a', 1), ('b', 2)]"]);
}

#[test]
fn test_types_method_type_binding() {
    let out = run_python(
        r#"
from types import MethodType

class Greeter:
    def __init__(self, name):
        self.name = name

def custom_greet(self):
    return f"Hello, {self.name}!"

g = Greeter("Bob")
g.greet = MethodType(custom_greet, g)
print(g.greet())
"#,
    );
    assert_eq!(out, vec!["Hello, Bob!"]);
}

#[test]
fn test_types_new_class_dynamic_creation() {
    let out = run_python(
        r#"
import types

def populate_cls(ns):
    ns["val"] = 42
    ns["get_val"] = lambda self: self.val

MyClass = types.new_class("MyClass", (object,), exec_body=populate_cls)
obj = MyClass()
print(obj.val)
print(obj.get_val())
"#,
    );
    assert_eq!(out, vec!["42", "42"]);
}

#[test]
fn test_types_resolve_bases_tuple() {
    let out = run_python(
        r#"
import types

class BaseA: pass
class BaseB: pass

resolved = types.resolve_bases((BaseA, BaseB))
print(resolved == (BaseA, BaseB))
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_types_function_type_inspection() {
    let out = run_python(
        r#"
from types import FunctionType
def dummy(): pass
print(isinstance(dummy, FunctionType))
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_types_lambda_is_function_type() {
    let out = run_python(
        r#"
from types import FunctionType
f = lambda x: x + 1
print(isinstance(f, FunctionType))
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_types_simple_namespace_dict_attribute() {
    let out = run_python(
        r#"
from types import SimpleNamespace
ns = SimpleNamespace(x=10, y=20)
print(ns.__dict__)
"#,
    );
    assert_eq!(out, vec!["{'x': 10, 'y': 20}"]);
}

#[test]
fn test_types_simple_namespace_kwargs_unpacking() {
    let out = run_python(
        r#"
from types import SimpleNamespace
data = {"a": 100, "b": 200}
ns = SimpleNamespace(**data)
print(ns.a, ns.b)
"#,
    );
    assert_eq!(out, vec!["100 200"]);
}

#[test]
fn test_types_mapping_proxy_type_copy() {
    let out = run_python(
        r#"
from types import MappingProxyType
d = {"x": 1}
proxy = MappingProxyType(d)
proxy_copy = proxy.copy()
print(proxy_copy["x"])
print(isinstance(proxy_copy, MappingProxyType))
"#,
    );
    assert_eq!(out, vec!["1", "True"]);
}

#[test]
fn test_types_mapping_proxy_type_get_fallback() {
    let out = run_python(
        r#"
from types import MappingProxyType
proxy = MappingProxyType({"existing": 1})
print(proxy.get("existing", 0))
print(proxy.get("non_existent", 42))
"#,
    );
    assert_eq!(out, vec!["1", "42"]);
}

#[test]
fn test_types_dynamic_class_attribute() {
    let out = run_python(
        r#"
from types import DynamicClassAttribute

class Config:
    def __init__(self, debug):
        self._debug = debug

    @DynamicClassAttribute
    def debug(self):
        return self._debug

c = Config(True)
print(c.debug)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_types_generator_type_check() {
    let out = run_python(
        r#"
from types import GeneratorType
def gen(): yield 1
g = gen()
print(isinstance(g, GeneratorType))
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_types_coroutine_type_check() {
    let out = run_python(
        r#"
from types import CoroutineType
async def coro(): pass
c = coro()
print(isinstance(c, CoroutineType))
c.close()
"#,
    );
    assert_eq!(out, vec!["True"]);
}
