use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: Tuples, NamedTuples & Unpacking — immutability, packing, unpacking, typing.NamedTuple, collections.namedtuple
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_tuple_immutability_and_single_element() {
    let src = r#"
single = (42,)
not_tuple = (42)

print(type(single).__name__)
print(type(not_tuple).__name__)

try:
    single[0] = 99
except TypeError:
    print("TypeError: tuple immutable")
"#;
    assert_eq!(
        run_python(src),
        vec!["tuple", "int", "TypeError: tuple immutable"]
    );
}

#[test]
fn test_py_collections_namedtuple_creation_access() {
    let src = r#"
from collections import namedtuple

Point = namedtuple("Point", ["x", "y"])
p = Point(10, 20)

print(p.x, p.y)
print(p[0], p[1])
print(isinstance(p, tuple))
"#;
    assert_eq!(run_python(src), vec!["10 20", "10 20", "True"]);
}

#[test]
fn test_py_collections_namedtuple_asdict_replace() {
    let src = r#"
from collections import namedtuple

User = namedtuple("User", ["id", "name", "role"], defaults=["guest"])
u1 = User(1, "Alice")
print(u1)
print(u1._asdict())

u2 = u1._replace(name="Alice Smith", role="admin")
print(u2)
"#;
    assert_eq!(
        run_python(src),
        vec![
            "User(id=1, name='Alice', role='guest')",
            "{'id': 1, 'name': 'Alice', 'role': 'guest'}",
            "User(id=1, name='Alice Smith', role='admin')"
        ]
    );
}

#[test]
fn test_py_typing_namedtuple_class_syntax() {
    let src = r#"
from typing import NamedTuple

class Employee(NamedTuple):
    id: int
    name: str
    department: str = "Engineering"

e = Employee(101, "Bob")
print(e.id, e.name, e.department)
print(e._fields)
"#;
    assert_eq!(
        run_python(src),
        vec!["101 Bob Engineering", "('id', 'name', 'department')"]
    );
}

#[test]
fn test_py_tuple_extended_unpacking_starred() {
    let src = r#"
first, *middle, last = (1, 2, 3, 4, 5)
print(first)
print(middle)
print(last)
"#;
    assert_eq!(run_python(src), vec!["1", "[2, 3, 4]", "5"]);
}

#[test]
fn test_py_nested_tuple_unpacking() {
    let src = r#"
data = ("Alice", (1995, 5, 12), "Engineer")
name, (year, month, day), job = data
print(name, year, job)
"#;
    assert_eq!(run_python(src), vec!["Alice 1995 Engineer"]);
}

#[test]
fn test_py_tuple_comparison_lexicographical() {
    let src = r#"
print((1, 2, 3) < (1, 2, 4))
print((1, 3) > (1, 2, 99))
print(("apple", 10) < ("banana", 1))
"#;
    assert_eq!(run_python(src), vec!["True", "True", "True"]);
}

#[test]
fn test_py_tuple_methods_count_index() {
    let src = r#"
t = (1, 2, 3, 2, 1, 2)
print(t.count(2))
print(t.index(3))
print(t.index(2, 2))  # search from index 2
"#;
    assert_eq!(run_python(src), vec!["3", "2", "3"]);
}

#[test]
fn test_py_tuple_concatenation_and_repetition() {
    let src = r#"
t1 = (1, 2)
t2 = (3, 4)
print(t1 + t2)
print(t1 * 3)
"#;
    assert_eq!(run_python(src), vec!["(1, 2, 3, 4)", "(1, 2, 1, 2, 1, 2)"]);
}

#[test]
fn test_py_swap_variables_via_tuple_packing() {
    let src = r#"
a = 10
b = 20
a, b = b, a
print(a, b)
"#;
    assert_eq!(run_python(src), vec!["20 10"]);
}
