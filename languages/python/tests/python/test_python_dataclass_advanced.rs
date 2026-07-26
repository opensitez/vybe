use super::helpers::run_python;

#[test]
fn test_python_dataclass_defaults_and_init() {
    let src = r#"
from dataclasses import dataclass, field

@dataclass
class P:
    x: int
    y: int = 2
    tags: list = field(default_factory=list)

p = P(1)
p.tags.append('a')
print(p.x)
print(p.y)
print(p.tags)
"#;
    assert_eq!(run_python(src), vec!["1", "2", "['a']"]);
}

#[test]
fn test_python_dataclass_frozen() {
    let src = r#"
from dataclasses import dataclass

@dataclass(frozen=True)
class Point:
    x: int
    y: int

p = Point(1, 2)
print(p.x + p.y)
"#;
    assert_eq!(run_python(src), vec!["3"]);
}
