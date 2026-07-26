// Python dataclass inheritance — parent/child, field ordering, post_init, super
use super::helpers::run_python;

#[test]
fn test_dataclass_inherits_fields() {
    let script = r#"
from dataclasses import dataclass

@dataclass
class Base:
    x: int
    y: int

@dataclass
class Child(Base):
    z: int

c = Child(1, 2, 3)
print(c.x, c.y, c.z)
"#;
    assert_eq!(run_python(script), vec!["1 2 3"]);
}

#[test]
fn test_dataclass_override_default() {
    let script = r#"
from dataclasses import dataclass

@dataclass
class Base:
    x: int = 10

@dataclass
class Child(Base):
    y: int = 20

c = Child()
print(c.x, c.y)
c2 = Child(x=100)
print(c2.x, c2.y)
"#;
    assert_eq!(run_python(script), vec!["10 20", "100 20"]);
}

#[test]
fn test_dataclass_post_init_chain() {
    let script = r#"
from dataclasses import dataclass

@dataclass
class Point:
    x: float
    y: float
    magnitude: float = 0.0

    def __post_init__(self):
        self.magnitude = (self.x**2 + self.y**2) ** 0.5

p = Point(3.0, 4.0)
print(p.magnitude)
"#;
    assert_eq!(run_python(script), vec!["5.0"]);
}

#[test]
fn test_dataclass_frozen_parent() {
    let script = r#"
from dataclasses import dataclass

@dataclass(frozen=True)
class Immutable:
    value: int

obj = Immutable(42)
print(obj.value)
try:
    obj.value = 99
    print("no_error")
except (AttributeError, TypeError, Exception):
    print("immutable")
"#;
    assert_eq!(run_python(script), vec!["42", "immutable"]);
}

#[test]
fn test_dataclass_equality() {
    let script = r#"
from dataclasses import dataclass

@dataclass
class Point:
    x: int
    y: int

p1 = Point(1, 2)
p2 = Point(1, 2)
p3 = Point(3, 4)
print(p1 == p2)
print(p1 == p3)
"#;
    assert_eq!(run_python(script), vec!["True", "False"]);
}

#[test]
fn test_dataclass_repr() {
    let script = r#"
from dataclasses import dataclass

@dataclass
class Color:
    r: int
    g: int
    b: int

c = Color(255, 0, 128)
print(repr(c))
"#;
    assert_eq!(run_python(script), vec!["Color(r=255, g=0, b=128)"]);
}

#[test]
fn test_dataclass_field_factory() {
    let script = r#"
from dataclasses import dataclass, field

@dataclass
class Container:
    items: list = field(default_factory=list)

c1 = Container()
c2 = Container()
c1.items.append(1)
print(c1.items)
print(c2.items)
"#;
    assert_eq!(run_python(script), vec!["[1]", "[]"]);
}
