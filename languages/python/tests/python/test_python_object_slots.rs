// Python __slots__ — memory optimization, slot access, no __dict__, inheritance
use super::helpers::run_python;

#[test]
fn test_slots_basic() {
    let script = r#"
class Point:
    __slots__ = ['x', 'y']
    def __init__(self, x, y):
        self.x = x
        self.y = y

p = Point(3, 4)
print(p.x, p.y)
"#;
    assert_eq!(run_python(script), vec!["3 4"]);
}

#[test]
fn test_slots_no_dict() {
    let script = r#"
class Compact:
    __slots__ = ('a', 'b')

c = Compact()
c.a = 1
c.b = 2
print(hasattr(c, '__dict__'))
"#;
    assert_eq!(run_python(script), vec!["False"]);
}

#[test]
fn test_slots_rejects_extra_attr() {
    let script = r#"
class Locked:
    __slots__ = ('x',)

obj = Locked()
obj.x = 10
try:
    obj.y = 20
    print("allowed")
except AttributeError:
    print("denied")
"#;
    assert_eq!(run_python(script), vec!["denied"]);
}

#[test]
fn test_slots_tuple_syntax() {
    let script = r#"
class Vec3:
    __slots__ = ('x', 'y', 'z')
    def __init__(self, x, y, z):
        self.x, self.y, self.z = x, y, z

v = Vec3(1, 2, 3)
print(v.x + v.y + v.z)
"#;
    assert_eq!(run_python(script), vec!["6"]);
}

#[test]
fn test_slots_in_subclass() {
    let script = r#"
class Base:
    __slots__ = ('x',)

class Child(Base):
    __slots__ = ('y',)

    def __init__(self, x, y):
        self.x = x
        self.y = y

c = Child(10, 20)
print(c.x, c.y)
"#;
    assert_eq!(run_python(script), vec!["10 20"]);
}

#[test]
fn test_slots_descriptor_access() {
    let script = r#"
class Counter:
    __slots__ = ('count',)
    def __init__(self):
        self.count = 0
    def inc(self):
        self.count += 1
        return self.count

c = Counter()
print(c.inc())
print(c.inc())
print(c.inc())
"#;
    assert_eq!(run_python(script), vec!["1", "2", "3"]);
}
