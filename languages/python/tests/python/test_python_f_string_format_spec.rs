// Python f-string format specs — !r, !s, !a, width, fill, align, sign, precision
use super::helpers::run_python;

#[test]
fn test_fstring_width_align() {
    let script = r#"
name = "hi"
print(f"{name:>10}")
print(f"{name:<10}|")
print(f"{name:^10}|")
"#;
    assert_eq!(run_python(script), vec!["        hi", "hi        |", "    hi    |"]);
}

#[test]
fn test_fstring_fill_char() {
    let script = r#"
n = 42
print(f"{n:0>5}")
print(f"{n:*<8}")
"#;
    assert_eq!(run_python(script), vec!["00042", "42******"]);
}

#[test]
fn test_fstring_number_precision() {
    let script = r#"
pi = 3.14159265
print(f"{pi:.2f}")
print(f"{pi:.4f}")
print(f"{pi:.0f}")
"#;
    assert_eq!(run_python(script), vec!["3.14", "3.1416", "3"]);
}

#[test]
fn test_fstring_sign_format() {
    let script = r#"
pos = 42
neg = -42
print(f"{pos:+}")
print(f"{neg:+}")
print(f"{pos: }")
"#;
    assert_eq!(run_python(script), vec!["+42", "-42", " 42"]);
}

#[test]
fn test_fstring_hex_oct_bin() {
    let script = r#"
n = 255
print(f"{n:x}")
print(f"{n:X}")
print(f"{n:o}")
print(f"{n:b}")
print(f"{n:#x}")
"#;
    assert_eq!(run_python(script), vec!["ff", "FF", "377", "11111111", "0xff"]);
}

#[test]
fn test_fstring_bang_r() {
    let script = r#"
s = "hello\nworld"
print(f"{s!r}")
"#;
    assert_eq!(run_python(script), vec!["'hello\\nworld'"]);
}

#[test]
fn test_fstring_bang_s() {
    let script = r#"
class Obj:
    def __str__(self):
        return "string"
    def __repr__(self):
        return "repr"

o = Obj()
print(f"{o!s}")
print(f"{o!r}")
"#;
    assert_eq!(run_python(script), vec!["string", "repr"]);
}

#[test]
fn test_fstring_nested_expression() {
    let script = r#"
width = 8
value = 3.14
print(f"{value:{width}.2f}")
"#;
    assert_eq!(run_python(script), vec!["    3.14"]);
}

#[test]
fn test_fstring_thousands_separator() {
    let script = r#"
n = 1234567
print(f"{n:,}")
print(f"{n:_}")
"#;
    assert_eq!(run_python(script), vec!["1,234,567", "1_234_567"]);
}
