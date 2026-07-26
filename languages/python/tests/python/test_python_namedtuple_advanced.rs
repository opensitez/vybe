// Python namedtuple advanced — _fields, _asdict, _make, _replace, defaults
use super::helpers::run_python;

#[test]
fn test_namedtuple_basic() {
    let script = r#"
from collections import namedtuple
Point = namedtuple('Point', ['x', 'y'])
p = Point(3, 4)
print(p.x, p.y)
print(p[0], p[1])
"#;
    assert_eq!(run_python(script), vec!["3 4", "3 4"]);
}

#[test]
fn test_namedtuple_fields() {
    let script = r#"
from collections import namedtuple
Color = namedtuple('Color', 'r g b')
print(Color._fields)
"#;
    assert_eq!(run_python(script), vec!["('r', 'g', 'b')"]);
}

#[test]
fn test_namedtuple_asdict() {
    let script = r#"
from collections import namedtuple
Person = namedtuple('Person', ['name', 'age'])
p = Person('Alice', 30)
d = p._asdict()
print(d['name'])
print(d['age'])
"#;
    assert_eq!(run_python(script), vec!["Alice", "30"]);
}

#[test]
fn test_namedtuple_make() {
    let script = r#"
from collections import namedtuple
Point = namedtuple('Point', ['x', 'y'])
p = Point._make([10, 20])
print(p)
"#;
    assert_eq!(run_python(script), vec!["Point(x=10, y=20)"]);
}

#[test]
fn test_namedtuple_replace() {
    let script = r#"
from collections import namedtuple
Point = namedtuple('Point', ['x', 'y'])
p1 = Point(1, 2)
p2 = p1._replace(y=99)
print(p1)
print(p2)
"#;
    assert_eq!(run_python(script), vec!["Point(x=1, y=2)", "Point(x=1, y=99)"]);
}

#[test]
fn test_namedtuple_defaults() {
    let script = r#"
from collections import namedtuple
Config = namedtuple('Config', ['host', 'port', 'debug'], defaults=['localhost', 8080, False])
c = Config()
print(c.host, c.port, c.debug)
c2 = Config('example.com')
print(c2.host, c2.port)
"#;
    assert_eq!(run_python(script), vec!["localhost 8080 False", "example.com 8080"]);
}

#[test]
fn test_namedtuple_is_tuple() {
    let script = r#"
from collections import namedtuple
Point = namedtuple('Point', ['x', 'y'])
p = Point(1, 2)
print(isinstance(p, tuple))
print(p == (1, 2))
"#;
    assert_eq!(run_python(script), vec!["True", "True"]);
}

#[test]
fn test_namedtuple_immutable() {
    let script = r#"
from collections import namedtuple
Point = namedtuple('Point', ['x', 'y'])
p = Point(1, 2)
try:
    p.x = 99
    print("mutable")
except AttributeError:
    print("immutable")
"#;
    assert_eq!(run_python(script), vec!["immutable"]);
}
