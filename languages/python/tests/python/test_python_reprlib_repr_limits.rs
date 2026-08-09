use super::helpers::run_python;

// reprlib — Repr class, maxstring/maxlist/maxdict/maxset/maxlong/maxother, recursive structures

#[test]
fn test_reprlib_default_repr_long_list_truncated() {
    let out = run_python(
        r#"
import reprlib
r = reprlib.Repr()
result = r.repr(list(range(100)))
print("..." in result)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_reprlib_maxlist_controls_list_length() {
    let out = run_python(
        r#"
import reprlib
r = reprlib.Repr()
r.maxlist = 3
result = r.repr(list(range(10)))
print("..." in result)
# Should have at most 3 elements shown
parts = result.strip("[]").split(", ")
print(len(parts) <= 4)  # 3 items + "..."
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_reprlib_maxstring_truncates_long_string() {
    let out = run_python(
        r#"
import reprlib
r = reprlib.Repr()
r.maxstring = 10
result = r.repr("a" * 50)
print("..." in result)
print(len(result) < 30)
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_reprlib_maxdict_limits_dict_entries() {
    let out = run_python(
        r#"
import reprlib
r = reprlib.Repr()
r.maxdict = 2
d = {i: i*2 for i in range(10)}
result = r.repr(d)
print("..." in result)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_reprlib_maxset_limits_set_entries() {
    let out = run_python(
        r#"
import reprlib
r = reprlib.Repr()
r.maxset = 2
result = r.repr(set(range(10)))
print("..." in result)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_reprlib_maxtuple_limits_tuple() {
    let out = run_python(
        r#"
import reprlib
r = reprlib.Repr()
r.maxtuple = 2
result = r.repr(tuple(range(10)))
print("..." in result)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_reprlib_maxlong_limits_large_int() {
    let out = run_python(
        r#"
import reprlib
r = reprlib.Repr()
r.maxlong = 10
big = 10**100
result = r.repr(big)
print("..." in result)
print(len(result) < 20)
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_reprlib_repr_recursive_list() {
    let out = run_python(
        r#"
import reprlib
lst = []
lst.append(lst)
result = reprlib.repr(lst)
print("[...]" in result)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_reprlib_repr_recursive_dict() {
    let out = run_python(
        r#"
import reprlib
d = {}
d["self"] = d
result = reprlib.repr(d)
print("{...}" in result)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_reprlib_maxother_limits_generic_repr() {
    let out = run_python(
        r#"
import reprlib
r = reprlib.Repr()
r.maxother = 10
class Big:
    def __repr__(self):
        return "x" * 200
result = r.repr(Big())
print(len(result) < 20)
print("..." in result)
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_reprlib_module_level_repr_function() {
    let out = run_python(
        r#"
import reprlib
# Module-level reprlib.repr() uses default Repr
result = reprlib.repr(list(range(200)))
print("..." in result)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_reprlib_short_string_not_truncated() {
    let out = run_python(
        r#"
import reprlib
r = reprlib.Repr()
r.maxstring = 30
result = r.repr("hello")
print("..." not in result)
print("hello" in result)
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_reprlib_small_list_not_truncated() {
    let out = run_python(
        r#"
import reprlib
r = reprlib.Repr()
r.maxlist = 10
result = r.repr([1, 2, 3])
print("..." not in result)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_reprlib_maxlevel_controls_nesting() {
    let out = run_python(
        r#"
import reprlib
r = reprlib.Repr()
r.maxlevel = 2
nested = [[[[1, 2, 3]]]]
result = r.repr(nested)
# At depth > maxlevel, inner content is truncated
print("..." in result)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_reprlib_maxfrozenset_limits() {
    let out = run_python(
        r#"
import reprlib
r = reprlib.Repr()
r.maxfrozenset = 2
result = r.repr(frozenset(range(10)))
print("..." in result)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_reprlib_repr_custom_class() {
    let out = run_python(
        r#"
import reprlib
class Point:
    def __init__(self, x, y):
        self.x = x
        self.y = y
    def __repr__(self):
        return f"Point({self.x}, {self.y})"
r = reprlib.Repr()
result = r.repr(Point(3, 4))
print("Point" in result)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_reprlib_recursive_depth_prevents_infinite_loop() {
    let out = run_python(
        r#"
import reprlib
# Deeply nested structure
lst = [None]
for _ in range(50):
    lst = [lst]
result = reprlib.repr(lst)
print(isinstance(result, str))
print("..." in result)
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_reprlib_maxdeque_limits() {
    let out = run_python(
        r#"
import reprlib
from collections import deque
r = reprlib.Repr()
r.maxdeque = 2
result = r.repr(deque(range(10)))
print("..." in result)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_reprlib_repr_bytes_object() {
    let out = run_python(
        r#"
import reprlib
r = reprlib.Repr()
r.maxstring = 5
result = r.repr(b"hello world")
print(isinstance(result, str))
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_reprlib_repr_nested_dict_truncates_inner() {
    let out = run_python(
        r#"
import reprlib
r = reprlib.Repr()
r.maxdict = 2
r.maxlevel = 1
d = {"a": list(range(100)), "b": list(range(100)), "c": list(range(100))}
result = r.repr(d)
print("..." in result)
"#,
    );
    assert_eq!(out, vec!["True"]);
}
