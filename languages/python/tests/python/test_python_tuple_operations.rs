// Python tuple operations — packing, unpacking, methods, named tuples, hashing
use super::helpers::run_python;

#[test]
fn test_tuple_packing() {
    let script = r#"
t = 1, 2, 3
print(type(t).__name__)
print(t)
"#;
    assert_eq!(run_python(script), vec!["tuple", "(1, 2, 3)"]);
}

#[test]
fn test_tuple_index_count() {
    let script = r#"
t = (1, 2, 3, 2, 1)
print(t.index(2))
print(t.count(1))
"#;
    assert_eq!(run_python(script), vec!["1", "2"]);
}

#[test]
fn test_tuple_concatenation() {
    let script = r#"
a = (1, 2)
b = (3, 4)
c = a + b
print(c)
print(len(c))
"#;
    assert_eq!(run_python(script), vec!["(1, 2, 3, 4)", "4"]);
}

#[test]
fn test_tuple_repetition() {
    let script = r#"
t = (0,) * 4
print(t)
"#;
    assert_eq!(run_python(script), vec!["(0, 0, 0, 0)"]);
}

#[test]
fn test_tuple_hashable() {
    let script = r#"
t1 = (1, 2, 3)
t2 = (1, 2, 3)
d = {t1: "value"}
print(d[t2])
print(hash(t1) == hash(t2))
"#;
    assert_eq!(run_python(script), vec!["value", "True"]);
}

#[test]
fn test_tuple_single_element() {
    let script = r#"
t = (42,)
print(type(t).__name__)
not_tuple = (42)
print(type(not_tuple).__name__)
"#;
    assert_eq!(run_python(script), vec!["tuple", "int"]);
}

#[test]
fn test_tuple_immutable() {
    let script = r#"
t = (1, 2, 3)
try:
    t[0] = 99
    print("mutable")
except TypeError:
    print("immutable")
"#;
    assert_eq!(run_python(script), vec!["immutable"]);
}

#[test]
fn test_tuple_min_max_sum() {
    let script = r#"
t = (3, 1, 4, 1, 5, 9, 2, 6)
print(min(t))
print(max(t))
print(sum(t))
"#;
    assert_eq!(run_python(script), vec!["1", "9", "31"]);
}

#[test]
fn test_tuple_comparison() {
    let script = r#"
print((1, 2, 3) < (1, 2, 4))
print((1, 2) < (1, 2, 0))
print((3,) > (2, 9))
"#;
    assert_eq!(run_python(script), vec!["True", "True", "True"]);
}
