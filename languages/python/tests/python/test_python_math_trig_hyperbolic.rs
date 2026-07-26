use super::helpers::run_python;

#[test]
fn test_python_math_hyperbolic_roundtrip() {
    let src = r#"
import math
x = 1.2
print(round(math.sinh(math.asinh(x)), 10))
print(round(math.cosh(1.0), 10))
"#;
    assert_eq!(run_python(src), vec!["1.2", "1.5430806348"]);
}

#[test]
fn test_python_math_tanh() {
    let src = r#"
import math
v = math.tanh(0.0)
print(v == 0.0)
"#;
    assert_eq!(run_python(src), vec!["True"]);
}
