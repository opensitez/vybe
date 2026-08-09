// Python conditional expressions — ternary, short-circuit, truthiness
use super::helpers::run_python;

#[test]
fn test_ternary_basic() {
    let script = r#"
x = 10
result = "big" if x > 5 else "small"
print(result)
"#;
    assert_eq!(run_python(script), vec!["big"]);
}

#[test]
fn test_ternary_nested() {
    let script = r#"
x = 5
result = "pos" if x > 0 else ("neg" if x < 0 else "zero")
print(result)
"#;
    assert_eq!(run_python(script), vec!["pos"]);
}

#[test]
fn test_short_circuit_and() {
    let script = r#"
calls = []
def f(v):
    calls.append(v)
    return v

result = f(False) and f(True)
print(result)
print(calls)
"#;
    assert_eq!(run_python(script), vec!["False", "[False]"]);
}

#[test]
fn test_short_circuit_or() {
    let script = r#"
calls = []
def f(v):
    calls.append(v)
    return v

result = f(True) or f(False)
print(result)
print(calls)
"#;
    assert_eq!(run_python(script), vec!["True", "[True]"]);
}

#[test]
fn test_and_returns_first_falsy() {
    let script = r#"
print(0 and 1)
print([] and "hello")
print(None and "x")
"#;
    assert_eq!(run_python(script), vec!["0", "[]", "None"]);
}

#[test]
fn test_or_returns_first_truthy() {
    let script = r#"
print(0 or 42)
print("" or "fallback")
print(None or [1, 2])
"#;
    assert_eq!(run_python(script), vec!["42", "fallback", "[1, 2]"]);
}

#[test]
fn test_ternary_in_list_comp() {
    let script = r#"
data = [1, -2, 3, -4, 5]
signs = ["pos" if x > 0 else "neg" for x in data]
print(signs)
"#;
    assert_eq!(
        run_python(script),
        vec!["['pos', 'neg', 'pos', 'neg', 'pos']"]
    );
}

#[test]
fn test_conditional_assignment_chained() {
    let script = r#"
a = b = c = None
a = a or b or c or "default"
print(a)
"#;
    assert_eq!(run_python(script), vec!["default"]);
}

#[test]
fn test_not_operator() {
    let script = r#"
print(not True)
print(not False)
print(not 0)
print(not "hello")
print(not "")
"#;
    assert_eq!(
        run_python(script),
        vec!["False", "True", "True", "False", "True"]
    );
}
