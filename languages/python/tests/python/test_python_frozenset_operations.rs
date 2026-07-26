// Python frozenset — creation, operations, hashability, set algebra
use super::helpers::run_python;

#[test]
fn test_frozenset_creation() {
    let script = r#"
fs = frozenset([1, 2, 3, 2, 1])
print(len(fs))
print(2 in fs)
"#;
    assert_eq!(run_python(script), vec!["3", "True"]);
}

#[test]
fn test_frozenset_as_dict_key() {
    let script = r#"
d = {}
fs = frozenset([1, 2, 3])
d[fs] = "triangle"
print(d[frozenset([3, 2, 1])])
"#;
    assert_eq!(run_python(script), vec!["triangle"]);
}

#[test]
fn test_frozenset_as_set_element() {
    let script = r#"
s = set()
fs = frozenset([1, 2])
s.add(fs)
print(fs in s)
print(len(s))
"#;
    assert_eq!(run_python(script), vec!["True", "1"]);
}

#[test]
fn test_frozenset_union_intersection() {
    let script = r#"
a = frozenset([1, 2, 3])
b = frozenset([2, 3, 4])
print(sorted(a | b))
print(sorted(a & b))
print(sorted(a - b))
"#;
    assert_eq!(run_python(script), vec!["[1, 2, 3, 4]", "[2, 3]", "[1]"]);
}

#[test]
fn test_frozenset_is_subset_superset() {
    let script = r#"
a = frozenset([1, 2])
b = frozenset([1, 2, 3])
print(a.issubset(b))
print(b.issuperset(a))
print(a.isdisjoint(frozenset([4, 5])))
"#;
    assert_eq!(run_python(script), vec!["True", "True", "True"]);
}

#[test]
fn test_frozenset_immutable() {
    let script = r#"
fs = frozenset([1, 2])
try:
    fs.add(3)
    print("no_error")
except AttributeError:
    print("AttributeError")
"#;
    assert_eq!(run_python(script), vec!["AttributeError"]);
}

#[test]
fn test_frozenset_symmetric_difference() {
    let script = r#"
a = frozenset([1, 2, 3])
b = frozenset([2, 3, 4])
print(sorted(a ^ b))
"#;
    assert_eq!(run_python(script), vec!["[1, 4]"]);
}

#[test]
fn test_frozenset_copy() {
    let script = r#"
a = frozenset([1, 2, 3])
b = a.copy()
print(a == b)
print(a is b)
"#;
    assert_eq!(run_python(script), vec!["True", "False"]);
}
