use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: Dataclass Fields & Config — @dataclass, field(), default_factory, frozen, post_init, kw_only
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_dataclass_basic_field_generation() {
    let src = r#"
from dataclasses import dataclass

@dataclass
class Point:
    x: float
    y: float

p = Point(1.5, 2.5)
print(p.x, p.y)
print(repr(p))
"#;
    assert_eq!(run_python(src), vec!["1.5 2.5", "Point(x=1.5, y=2.5)"]);
}

#[test]
fn test_py_dataclass_default_factory_mutable_fields() {
    let src = r#"
from dataclasses import dataclass, field
from typing import List

@dataclass
class Group:
    name: str
    members: List[str] = field(default_factory=list)

g1 = Group("Admin")
g2 = Group("Users")
g1.members.append("Alice")
print(g1.members)
print(g2.members)  # independent lists
"#;
    assert_eq!(run_python(src), vec!["['Alice']", "[]"]);
}

#[test]
fn test_py_dataclass_post_init_validation() {
    let src = r#"
from dataclasses import dataclass

@dataclass
class Rectangle:
    width: float
    height: float

    def __post_init__(self):
        if self.width <= 0 or self.height <= 0:
            raise ValueError("Dimensions must be positive")
        self.area = self.width * self.height

r = Rectangle(4.0, 5.0)
print(r.area)
try:
    Rectangle(-1.0, 5.0)
except ValueError as e:
    print(e)
"#;
    assert_eq!(run_python(src), vec!["20.0", "Dimensions must be positive"]);
}

#[test]
fn test_py_dataclass_frozen_immutability() {
    let src = r#"
from dataclasses import dataclass

@dataclass(frozen=True)
class Config:
    host: str
    port: int

cfg = Config("localhost", 8080)
print(cfg.host, cfg.port)
try:
    cfg.port = 9090
except Exception as e:
    print(type(e).__name__)
"#;
    assert_eq!(
        run_python(src),
        vec!["localhost 8080", "FrozenInstanceError"]
    );
}

#[test]
fn test_py_dataclass_kw_only_fields() {
    let src = r#"
from dataclasses import dataclass

@dataclass(kw_only=True)
class Event:
    id: str
    payload: dict

e = Event(id="evt1", payload={"a": 1})
print(e.id, e.payload)
try:
    Event("evt1", {"a": 1})
except TypeError:
    print("TypeError: keyword-only required")
"#;
    assert_eq!(
        run_python(src),
        vec!["evt1 {'a': 1}", "TypeError: keyword-only required"]
    );
}

#[test]
fn test_py_dataclass_ordering_comparisons() {
    let src = r#"
from dataclasses import dataclass

@dataclass(order=True)
class Version:
    major: int
    minor: int
    patch: int

v1 = Version(1, 2, 0)
v2 = Version(1, 10, 0)
print(v1 < v2)
print(sorted([v2, v1]))
"#;
    assert_eq!(
        run_python(src),
        vec![
            "True",
            "[Version(major=1, minor=2, patch=0), Version(major=1, minor=10, patch=0)]"
        ]
    );
}

#[test]
fn test_py_dataclass_asdict_astuple_conversions() {
    let src = r#"
from dataclasses import dataclass, asdict, astuple

@dataclass
class User:
    name: str
    age: int

u = User("Bob", 30)
print(asdict(u))
print(astuple(u))
"#;
    assert_eq!(
        run_python(src),
        vec!["{'name': 'Bob', 'age': 30}", "('Bob', 30)"]
    );
}

#[test]
fn test_py_dataclass_replace_shallow_copy() {
    let src = r#"
from dataclasses import dataclass, replace

@dataclass(frozen=True)
class ServerConfig:
    host: str = "localhost"
    port: int = 8080
    debug: bool = False

c1 = ServerConfig()
c2 = replace(c1, port=9090, debug=True)
print(c1)
print(c2)
"#;
    assert_eq!(
        run_python(src),
        vec![
            "ServerConfig(host='localhost', port=8080, debug=False)",
            "ServerConfig(host='localhost', port=9090, debug=True)"
        ]
    );
}

#[test]
fn test_py_dataclass_field_metadata() {
    let src = r#"
from dataclasses import dataclass, field, fields

@dataclass
class Schema:
    username: str = field(metadata={"description": "User login identifier"})

f = fields(Schema)[0]
print(f.name)
print(f.metadata["description"])
"#;
    assert_eq!(run_python(src), vec!["username", "User login identifier"]);
}

#[test]
fn test_py_dataclass_inheritance_field_ordering() {
    let src = r#"
from dataclasses import dataclass

@dataclass
class Base:
    x: int

@dataclass
class Child(Base):
    y: int

c = Child(1, 2)
print(c.x, c.y)
"#;
    assert_eq!(run_python(src), vec!["1 2"]);
}
