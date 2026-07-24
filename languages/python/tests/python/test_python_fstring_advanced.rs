use super::helpers::run_python;

// Language features: f-string = specifier, nested braces, !r !s !a with format spec

#[test]
fn test_fstring_equals_specifier_basic() {
    let out = run_python(r#"
x = 42
print(f"{x=}")
"#);
    assert_eq!(out, vec!["x=42"]);
}

#[test]
fn test_fstring_equals_specifier_with_expression() {
    let out = run_python(r#"
a, b = 3, 4
print(f"{a + b=}")
"#);
    assert_eq!(out, vec!["a + b=7"]);
}

#[test]
fn test_fstring_equals_specifier_with_format() {
    let out = run_python(r#"
x = 3.14159
print(f"{x=:.2f}")
"#);
    assert_eq!(out, vec!["x=3.14"]);
}

#[test]
fn test_fstring_conversion_r_uses_repr() {
    let out = run_python(r#"
s = "hello\nworld"
print(f"{s!r}")
"#);
    assert_eq!(out, vec!["'hello\\nworld'"]);
}

#[test]
fn test_fstring_conversion_s_uses_str() {
    let out = run_python(r#"
class Obj:
    def __str__(self): return "str_form"
    def __repr__(self): return "repr_form"
o = Obj()
print(f"{o!s}")
"#);
    assert_eq!(out, vec!["str_form"]);
}

#[test]
fn test_fstring_conversion_a_uses_ascii() {
    let out = run_python(r#"
s = "caf\u00e9"
print(f"{s!a}")
"#);
    assert_eq!(out, vec!["'caf\\xe9'"]);
}

#[test]
fn test_fstring_conversion_r_with_format_spec() {
    let out = run_python(r#"
s = "hi"
print(f"{s!r:^10}")
"#);
    assert_eq!(out, vec!["   'hi'   "]);
}

#[test]
fn test_fstring_nested_braces_expression() {
    let out = run_python(r#"
width = 10
print(f"{'hello':^{width}}")
"#);
    assert_eq!(out, vec!["  hello   "]);
}

#[test]
fn test_fstring_nested_format_spec_computed() {
    let out = run_python(r#"
precision = 3
value = 3.141592653589793
print(f"{value:.{precision}f}")
"#);
    assert_eq!(out, vec!["3.142"]);
}

#[test]
fn test_fstring_nested_width_and_fill() {
    let out = run_python(r#"
fill = "*"
width = 8
text = "hi"
print(f"{text:{fill}^{width}}")
"#);
    assert_eq!(out, vec!["***hi***"]);
}

#[test]
fn test_fstring_double_brace_literal() {
    let out = run_python(r#"
print(f"{{not interpolated}}")
"#);
    assert_eq!(out, vec!["{not interpolated}"]);
}

#[test]
fn test_fstring_multiline_expression() {
    let out = run_python(r#"
items = [1, 2, 3]
print(f"{sum(items)=}")
"#);
    assert_eq!(out, vec!["sum(items)=6"]);
}

#[test]
fn test_fstring_dict_access_in_expression() {
    let out = run_python(r#"
d = {"key": "value"}
print(f"result: {d['key']}")
"#);
    assert_eq!(out, vec!["result: value"]);
}

#[test]
fn test_fstring_conditional_expression() {
    let out = run_python(r#"
x = 5
print(f"{'even' if x % 2 == 0 else 'odd'}")
"#);
    assert_eq!(out, vec!["odd"]);
}

#[test]
fn test_fstring_lambda_in_expression() {
    let out = run_python(r#"
f_str = f"{(lambda x: x * 2)(5)}"
print(f_str)
"#);
    assert_eq!(out, vec!["10"]);
}

#[test]
fn test_fstring_nested_f_string() {
    let out = run_python(r#"
inner = "world"
outer = f"hello {f'{inner}'}"
print(outer)
"#);
    assert_eq!(out, vec!["hello world"]);
}

#[test]
fn test_fstring_format_spec_zero_pad() {
    let out = run_python(r#"
n = 42
print(f"{n:05d}")
"#);
    assert_eq!(out, vec!["00042"]);
}

#[test]
fn test_fstring_format_spec_hex() {
    let out = run_python(r#"
n = 255
print(f"{n:#010x}")
"#);
    assert_eq!(out, vec!["0x000000ff"]);
}

#[test]
fn test_fstring_format_spec_plus_sign() {
    let out = run_python(r#"
print(f"{42:+d}")
print(f"{-42:+d}")
"#);
    assert_eq!(out, vec!["+42", "-42"]);
}

#[test]
fn test_fstring_equals_specifier_string() {
    let out = run_python(r#"
name = "Alice"
print(f"{name=}")
"#);
    assert_eq!(out, vec!["name='Alice'"]);
}
