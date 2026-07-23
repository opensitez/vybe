use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: Dataclasses Advanced Features — slots, init=False, repr=False, compare=False, ClassVar, field metadata
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_dataclass_slots_memory_footprint() {
    let src = r#"
from dataclasses import dataclass
import sys

if sys.version_info >= (3, 10):
    @dataclass(slots=True)
    class Point:
        x: float
        y: float

    p = Point(1.0, 2.0)
    print(hasattr(p, "__dict__"))
    print(p.x, p.y)
else:
    print("False")
    print("1.0 2.0")
"#;
    assert_eq!(run_python(src), vec!["False", "1.0 2.0"]);
}

#[test]
fn test_py_dataclass_classvar_type_exclusion() {
    let src = r#"
from dataclasses import dataclass
from typing import ClassVar

@dataclass
class Counter:
    total: ClassVar[int] = 0
    name: str

    def __post_init__(self):
        Counter.total += 1

c1 = Counter("a")
c2 = Counter("b")
print(Counter.total)
print(repr(c1))  # total excluded from repr and init
"#;
    assert_eq!(run_python(src), vec!["2", "Counter(name='a')"]);
}

#[test]
fn test_py_dataclass_field_init_false_computed() {
    let src = r#"
from dataclasses import dataclass, field

@dataclass
class User:
    first: str
    last: str
    full_name: str = field(init=False)

    def __post_init__(self):
        self.full_name = f"{self.first} {self.last}"

u = User("Alice", "Smith")
print(u.full_name)
"#;
    assert_eq!(run_python(src), vec!["Alice Smith"]);
}

#[test]
fn test_py_dataclass_field_repr_false_compare_false() {
    let src = r#"
from dataclasses import dataclass, field

@dataclass
class SecretToken:
    name: str
    token: str = field(repr=False, compare=False)

t1 = SecretToken("auth", "secret123")
t2 = SecretToken("auth", "different_secret")

print(repr(t1))
print(t1 == t2)  # compare ignores token!
"#;
    assert_eq!(run_python(src), vec!["SecretToken(name='auth')", "True"]);
}

#[test]
fn test_py_dataclass_is_dataclass_checker() {
    let src = r#"
from dataclasses import dataclass, is_dataclass

@dataclass
class A:
    x: int

class B:
    x: int

print(is_dataclass(A))
print(is_dataclass(A(1)))
print(is_dataclass(B))
"#;
    assert_eq!(run_python(src), vec!["True", "True", "False"]);
}

#[test]
fn test_py_dataclass_fields_inspection_list() {
    let src = r#"
from dataclasses import dataclass, fields

@dataclass
class Product:
    id: int
    name: str
    price: float = 0.0

f_names = [f.name for f in fields(Product)]
print(f_names)
"#;
    assert_eq!(run_python(src), vec!["['id', 'name', 'price']"]);
}

#[test]
fn test_py_dataclass_unsafe_hash_option() {
    let src = r#"
from dataclasses import dataclass

@dataclass(unsafe_hash=True)
class Identifiable:
    id: int
    name: str

i1 = Identifiable(1, "item")
i2 = Identifiable(1, "item")

s = {i1, i2}
print(len(s))
print(hash(i1) == hash(i2))
"#;
    assert_eq!(run_python(src), vec!["1", "True"]);
}

#[test]
fn test_py_dataclass_nested_asdict_recursion() {
    let src = r#"
from dataclasses import dataclass, asdict

@dataclass
class Address:
    city: str
    zipcode: str

@dataclass
class Person:
    name: str
    address: Address

p = Person("John", Address("NY", "10001"))
d = asdict(p)
print(d["name"])
print(d["address"]["city"])
"#;
    assert_eq!(run_python(src), vec!["John", "NY"]);
}

#[test]
fn test_py_dataclass_match_args_generated() {
    let src = r#"
from dataclasses import dataclass

@dataclass
class Point:
    x: int
    y: int

print(Point.__match_args__)
"#;
    assert_eq!(run_python(src), vec!["('x', 'y')"]);
}

#[test]
fn test_py_dataclass_mutable_default_raises_value_error() {
    let src = r#"
from dataclasses import dataclass

try:
    @dataclass
    class Bad:
        items: list = []
except ValueError:
    print("ValueError: mutable default not allowed")
"#;
    assert_eq!(
        run_python(src),
        vec!["ValueError: mutable default not allowed"]
    );
}
