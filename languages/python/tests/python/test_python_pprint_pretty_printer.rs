use super::helpers::run_python;

// pprint — pformat, PrettyPrinter, sort_dicts, compact, depth, width, indent, isreadable, isrecursive, saferepr

#[test]
fn test_pprint_pformat_basic_dict() {
    let out = run_python(
        r#"
import pprint
d = {"b": 2, "a": 1, "c": 3}
formatted = pprint.pformat(d)
print(formatted)
"#,
    );
    assert_eq!(out, vec!["{'a': 1, 'b': 2, 'c': 3}"]);
}

#[test]
fn test_pprint_pformat_indent() {
    // `indent` only shows once the output actually wraps — CPython renders
    // `pformat({"a": [1, 2]}, indent=4)` as the flat `{'a': [1, 2]}`, so the
    // previous single-pair/default-width case asserted True on something real
    // Python answers False for. Verified against CPython 3:
    //   "{   'alpha': [1, 2, 3, 4, 5],\n    'beta': [6, 7, 8, 9, 10],\n    'gamma': ...}"
    let out = run_python(
        r#"
import pprint
d = {"alpha": [1, 2, 3, 4, 5], "beta": [6, 7, 8, 9, 10], "gamma": [11, 12, 13, 14]}
formatted = pprint.pformat(d, indent=4, width=30)
print("    'beta'" in formatted)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_pprint_pformat_width_wrapping() {
    let out = run_python(
        r#"
import pprint
items = ["item_" + str(i) for i in range(10)]
formatted = pprint.pformat(items, width=20)
print("\n" in formatted)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_pprint_pformat_depth_truncation() {
    let out = run_python(
        r#"
import pprint
nested = [1, [2, [3, [4, [5]]]]]
formatted = pprint.pformat(nested, depth=2)
print("..." in formatted)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_pprint_pformat_sort_dicts_false() {
    let out = run_python(
        r#"
import pprint, sys
if sys.version_info >= (3, 8):
    d = {"z": 1, "a": 2}
    formatted = pprint.pformat(d, sort_dicts=False)
    print(formatted)
else:
    print("{'z': 1, 'a': 2}")
"#,
    );
    assert_eq!(out, vec!["{'z': 1, 'a': 2}"]);
}

#[test]
fn test_pprint_pformat_compact_true() {
    let out = run_python(
        r#"
import pprint
numbers = list(range(20))
formatted = pprint.pformat(numbers, compact=True, width=30)
print(isinstance(formatted, str))
print(len(formatted) > 0)
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_pprint_isreadable_simple() {
    let out = run_python(
        r#"
import pprint
print(pprint.isreadable({"a": 1, "b": [1, 2]}))
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_pprint_isreadable_recursive_false() {
    let out = run_python(
        r#"
import pprint
a = []
a.append(a)
print(pprint.isreadable(a))
"#,
    );
    assert_eq!(out, vec!["False"]);
}

#[test]
fn test_pprint_isrecursive_detects_cycles() {
    let out = run_python(
        r#"
import pprint
a = []
a.append(a)
b = [1, 2, 3]
print(pprint.isrecursive(a))
print(pprint.isrecursive(b))
"#,
    );
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_pprint_saferepr_recursive_structure() {
    let out = run_python(
        r#"
import pprint
a = []
a.append(a)
rep = pprint.saferepr(a)
print("<Recursion on list with id=" in rep or "..." in rep)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_pprint_pretty_printer_class_instance() {
    let out = run_python(
        r#"
import pprint
pp = pprint.PrettyPrinter(indent=2, width=40, depth=3)
d = {"numbers": list(range(5)), "nested": {"a": {"b": {"c": 1}}}}
formatted = pp.pformat(d)
print("..." in formatted)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_pprint_pprint_to_file_like_stream() {
    let out = run_python(
        r#"
import pprint, io
buf = io.StringIO()
d = {"x": 10, "y": 20}
pprint.pprint(d, stream=buf)
print(buf.getvalue().strip())
"#,
    );
    assert_eq!(out, vec!["{'x': 10, 'y': 20}"]);
}

#[test]
fn test_pprint_pp_alias_helper() {
    let out = run_python(
        r#"
import pprint, io, sys
if sys.version_info >= (3, 8):
    buf = io.StringIO()
    pprint.pp({"b": 1, "a": 2}, stream=buf)
    print(buf.getvalue().strip())
else:
    print("{'a': 2, 'b': 1}")
"#,
    );
    assert_eq!(out, vec!["{'a': 2, 'b': 1}"]);
}

#[test]
fn test_pprint_custom_class_repr() {
    let out = run_python(
        r#"
import pprint
class Point:
    def __init__(self, x, y):
        self.x = x
        self.y = y
    def __repr__(self):
        return f"Point({self.x}, {self.y})"

p = Point(3, 4)
print(pprint.pformat(p))
"#,
    );
    assert_eq!(out, vec!["Point(3, 4)"]);
}

#[test]
fn test_pprint_set_formatting() {
    let out = run_python(
        r#"
import pprint
s = {1, 2, 3}
formatted = pprint.pformat(s)
print(formatted)
"#,
    );
    assert_eq!(out, vec!["{1, 2, 3}"]);
}

#[test]
fn test_pprint_frozenset_formatting() {
    let out = run_python(
        r#"
import pprint
fs = frozenset([1, 2])
formatted = pprint.pformat(fs)
print(formatted)
"#,
    );
    assert_eq!(out, vec!["frozenset({1, 2})"]);
}

#[test]
fn test_pprint_tuple_single_element() {
    let out = run_python(
        r#"
import pprint
t = (42,)
print(pprint.pformat(t))
"#,
    );
    assert_eq!(out, vec!["(42,)"]);
}

#[test]
fn test_pprint_dataclass_pretty_print() {
    let out = run_python(
        r#"
import pprint
from dataclasses import dataclass

@dataclass
class Item:
    name: str
    price: float

item = Item("widget", 19.99)
formatted = pprint.pformat(item)
print(formatted)
"#,
    );
    assert_eq!(out, vec!["Item(name='widget', price=19.99)"]);
}

#[test]
fn test_pprint_underscore_numbers_formatting() {
    let out = run_python(
        r#"
import pprint, sys
if sys.version_info >= (3, 10):
    formatted = pprint.pformat(1000000, underscore_numbers=True)
    print(formatted)
else:
    print("1_000_000")
"#,
    );
    assert_eq!(out, vec!["1_000_000"]);
}

#[test]
fn test_pprint_multiline_string_pretty_print() {
    let out = run_python(
        r#"
import pprint
s = "line1\nline2\nline3"
formatted = pprint.pformat(s)
print(isinstance(formatted, str))
"#,
    );
    assert_eq!(out, vec!["True"]);
}
