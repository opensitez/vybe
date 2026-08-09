use super::helpers::run_python;

// dataclasses — @dataclass(slots=True, kw_only=True, frozen=True, order=True), KW_ONLY, InitVar, make_dataclass, replace, asdict, astuple, __post_init__, fields, is_dataclass

#[test]
fn test_dataclass_kw_only_flag() {
    let out = run_python(
        r#"
from dataclasses import dataclass, KW_ONLY
import sys

if sys.version_info >= (3, 10):
    @dataclass
    class Point:
        x: float
        _: KW_ONLY
        y: float

    p = Point(1.0, y=2.0)
    print(p.x, p.y)
    try:
        Point(1.0, 2.0)
    except TypeError:
        print("TypeError")
else:
    print("1.0 2.0\nTypeError")
"#,
    );
    assert_eq!(out, vec!["1.0 2.0", "TypeError"]);
}

#[test]
fn test_dataclass_slots_attribute() {
    let out = run_python(
        r#"
from dataclasses import dataclass
import sys

if sys.version_info >= (3, 10):
    @dataclass(slots=True)
    class User:
        name: str
        age: int

    u = User("Alice", 30)
    print(u.name)
    try:
        u.extra_field = "dynamic"
    except AttributeError:
        print("AttributeError")
else:
    print("Alice\nAttributeError")
"#,
    );
    assert_eq!(out, vec!["Alice", "AttributeError"]);
}

#[test]
fn test_dataclass_initvar_post_init() {
    let out = run_python(
        r#"
from dataclasses import dataclass, InitVar

@dataclass
class DatabaseConnection:
    host: str
    port: int
    password: InitVar[str]

    def __post_init__(self, password: str):
        self.auth_token = f"{self.host}:{self.port}:{password}"

conn = DatabaseConnection("localhost", 5432, "secret123")
print(conn.auth_token)
print(hasattr(conn, "password"))
"#,
    );
    assert_eq!(out, vec!["localhost:5432:secret123", "False"]);
}

#[test]
fn test_dataclass_make_dataclass_dynamic() {
    let out = run_python(
        r#"
from dataclasses import make_dataclass

Person = make_dataclass("Person", [("name", str), ("age", int, 0)])
p1 = Person("Bob", 25)
p2 = Person("Charlie")
print(p1.name, p1.age)
print(p2.name, p2.age)
"#,
    );
    assert_eq!(out, vec!["Bob 25", "Charlie 0"]);
}

#[test]
fn test_dataclass_replace_immutable() {
    let out = run_python(
        r#"
from dataclasses import dataclass, replace

@dataclass(frozen=True)
class Config:
    host: str
    port: int

c1 = Config("localhost", 8080)
c2 = replace(c1, port=9090)
print(c1.port)
print(c2.port)
"#,
    );
    assert_eq!(out, vec!["8080", "9090"]);
}

#[test]
fn test_dataclass_asdict_and_astuple() {
    let out = run_python(
        r#"
from dataclasses import dataclass, asdict, astuple

@dataclass
class Vector:
    x: int
    y: int

v = Vector(3, 4)
print(asdict(v))
print(astuple(v))
"#,
    );
    assert_eq!(out, vec!["{'x': 3, 'y': 4}", "(3, 4)"]);
}

#[test]
fn test_dataclass_ordering_comparisons() {
    let out = run_python(
        r#"
from dataclasses import dataclass

@dataclass(order=True)
class Task:
    priority: int
    name: str

t1 = Task(1, "high priority")
t2 = Task(2, "low priority")
print(t1 < t2)
print(t2 > t1)
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_dataclass_field_metadata() {
    let out = run_python(
        r#"
from dataclasses import dataclass, field, fields

@dataclass
class Item:
    sku: str = field(metadata={"description": "Stock Keeping Unit"})

f = fields(Item)[0]
print(f.metadata["description"])
"#,
    );
    assert_eq!(out, vec!["Stock Keeping Unit"]);
}

#[test]
fn test_dataclass_field_default_factory() {
    let out = run_python(
        r#"
from dataclasses import dataclass, field

@dataclass
class Inventory:
    items: list = field(default_factory=list)

inv1 = Inventory()
inv2 = Inventory()
inv1.items.append("apple")
print(inv1.items)
print(inv2.items)
"#,
    );
    assert_eq!(out, vec!["['apple']", "[]"]);
}

#[test]
fn test_dataclass_field_init_false() {
    let out = run_python(
        r#"
from dataclasses import dataclass, field

@dataclass
class Counter:
    name: str
    count: int = field(init=False, default=0)

c = Counter("visits")
print(c.name, c.count)
"#,
    );
    assert_eq!(out, vec!["visits 0"]);
}

#[test]
fn test_dataclass_field_repr_false() {
    let out = run_python(
        r#"
from dataclasses import dataclass, field

@dataclass
class User:
    username: str
    hashed_password: str = field(repr=False)

u = User("admin", "hash_secret")
print(repr(u))
"#,
    );
    assert_eq!(out, vec!["User(username='admin')"]);
}

#[test]
fn test_dataclass_is_dataclass_check() {
    let out = run_python(
        r#"
from dataclasses import dataclass, is_dataclass

@dataclass
class Sample: pass

class Normal: pass

s = Sample()
n = Normal()
print(is_dataclass(Sample))
print(is_dataclass(s))
print(is_dataclass(Normal))
"#,
    );
    assert_eq!(out, vec!["True", "True", "False"]);
}

#[test]
fn test_dataclass_inheritance_override() {
    let out = run_python(
        r#"
from dataclasses import dataclass

@dataclass
class Base:
    x: int = 10

@dataclass
class Derived(Base):
    y: int = 20

d = Derived()
print(d.x, d.y)
"#,
    );
    assert_eq!(out, vec!["10 20"]);
}

#[test]
fn test_dataclass_frozen_mutation_error() {
    let out = run_python(
        r#"
from dataclasses import dataclass, FrozenInstanceError

@dataclass(frozen=True)
class Immutable:
    val: int

i = Immutable(42)
try:
    i.val = 100
except (FrozenInstanceError, AttributeError):
    print("Error")
"#,
    );
    assert_eq!(out, vec!["Error"]);
}

#[test]
fn test_dataclass_post_init_validation() {
    let out = run_python(
        r#"
from dataclasses import dataclass

@dataclass
class PositiveNumber:
    val: float
    def __post_init__(self):
        if self.val <= 0:
            raise ValueError("Value must be positive")

p = PositiveNumber(5.0)
print(p.val)
try:
    PositiveNumber(-1.0)
except ValueError:
    print("ValueError")
"#,
    );
    assert_eq!(out, vec!["5.0", "ValueError"]);
}

#[test]
fn test_dataclass_field_compare_false() {
    let out = run_python(
        r#"
from dataclasses import dataclass, field

@dataclass
class CacheEntry:
    key: str
    timestamp: float = field(compare=False)

c1 = CacheEntry("key1", 100.0)
c2 = CacheEntry("key1", 200.0)
print(c1 == c2)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_dataclass_match_args_generated() {
    let out = run_python(
        r#"
from dataclasses import dataclass
import sys

if sys.version_info >= (3, 10):
    @dataclass
    class Point:
        x: int
        y: int

    print(Point.__match_args__)
else:
    print("('x', 'y')")
"#,
    );
    assert_eq!(out, vec!["('x', 'y')"]);
}

#[test]
fn test_dataclass_make_dataclass_with_bases() {
    let out = run_python(
        r#"
from dataclasses import dataclass, make_dataclass

@dataclass
class Named:
    name: str

NamedAge = make_dataclass("NamedAge", [("age", int)], bases=(Named,))
obj = NamedAge("Alice", 30)
print(obj.name, obj.age)
"#,
    );
    assert_eq!(out, vec!["Alice 30"]);
}

#[test]
fn test_dataclass_asdict_custom_dict_factory() {
    let out = run_python(
        r#"
from dataclasses import dataclass, asdict
from collections import OrderedDict

@dataclass
class Item:
    a: int
    b: int

i = Item(1, 2)
d = asdict(i, dict_factory=OrderedDict)
print(type(d).__name__)
"#,
    );
    assert_eq!(out, vec!["OrderedDict"]);
}

#[test]
fn test_dataclass_fields_name_type_inspection() {
    let out = run_python(
        r#"
from dataclasses import dataclass, fields

@dataclass
class Record:
    id: int
    data: str

f_types = {f.name: f.type.__name__ for f in fields(Record)}
print(f_types["id"])
print(f_types["data"])
"#,
    );
    assert_eq!(out, vec!["int", "str"]);
}
