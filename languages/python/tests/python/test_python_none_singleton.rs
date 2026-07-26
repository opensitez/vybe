// Python None singleton — identity, comparisons, truthiness, default params
use super::helpers::run_python;

#[test]
fn test_none_is_singleton() {
    let script = r#"
a = None
b = None
print(a is b)
print(a is None)
"#;
    assert_eq!(run_python(script), vec!["True", "True"]);
}

#[test]
fn test_none_equality_vs_identity() {
    let script = r#"
class NoneImpostor:
    def __eq__(self, other):
        return True

obj = NoneImpostor()
print(obj == None)
print(obj is None)
"#;
    assert_eq!(run_python(script), vec!["True", "False"]);
}

#[test]
fn test_none_is_falsy() {
    let script = r#"
print(bool(None))
print(not None)
"#;
    assert_eq!(run_python(script), vec!["False", "True"]);
}

#[test]
fn test_none_as_default_param_sentinel() {
    let script = r#"
def func(items=None):
    if items is None:
        items = []
    items.append(1)
    return items

print(func())
print(func())
print(func([10]))
"#;
    assert_eq!(run_python(script), vec!["[1]", "[1]", "[10, 1]"]);
}

#[test]
fn test_none_type() {
    let script = r#"
print(type(None).__name__)
print(isinstance(None, type(None)))
"#;
    assert_eq!(run_python(script), vec!["NoneType", "True"]);
}

#[test]
fn test_none_in_collections() {
    let script = r#"
lst = [1, None, 2, None, 3]
count = lst.count(None)
print(count)
non_none = [x for x in lst if x is not None]
print(non_none)
"#;
    assert_eq!(run_python(script), vec!["2", "[1, 2, 3]"]);
}

#[test]
fn test_function_returns_none_implicitly() {
    let script = r#"
def nothing():
    pass

result = nothing()
print(result is None)
print(result)
"#;
    assert_eq!(run_python(script), vec!["True", "None"]);
}
