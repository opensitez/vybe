use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: String Formatting & Specifiers — f-strings, __format__, specifiers, fill, alignment, precision
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_fstring_alignment_and_padding() {
    let src = r#"
s = "test"
print(f"{s:>10}")
print(f"{s:<10}")
print(f"{s:^10}")
print(f"{s:*^10}")
"#;
    assert_eq!(
        run_python(src),
        vec!["      test", "test      ", "   test   ", "***test***"]
    );
}

#[test]
fn test_py_fstring_number_formatting() {
    let src = r#"
n = 1234567.8910
print(f"{n:,.2f}")
print(f"{n:_>15,.2f}")
val = 42
print(f"{val:05d}")
print(f"{val:#x}")
print(f"{val:#010b}")
"#;
    assert_eq!(
        run_python(src),
        vec![
            "1,234,567.89",
            "___1,234,567.89",
            "00042",
            "0x2a",
            "0b00101010"
        ]
    );
}

#[test]
fn test_py_fstring_percentage_and_exponent() {
    let src = r#"
ratio = 0.756
print(f"{ratio:.1%}")
print(f"{ratio:.2%}")
num = 12345.678
print(f"{num:.2e}")
"#;
    assert_eq!(run_python(src), vec!["75.6%", "75.60%", "1.23e+04"]);
}

#[test]
fn test_py_fstring_nested_expressions() {
    let src = r#"
width = 10
precision = 3
val = 12.34567
print(f"{val:{width}.{precision}f}")
cols = ["Name", "Score"]
print(f"{cols[0]:<8} | {cols[1]:>8}")
"#;
    assert_eq!(run_python(src), vec!["    12.346", "Name     |    Score"]);
}

#[test]
fn test_py_fstring_debugging_equal_sign() {
    let src = r#"
x = 10
y = 25
print(f"{x=}")
print(f"{x + y=}")
print(f"{x * 2 = }")
"#;
    assert_eq!(run_python(src), vec!["x=10", "x + y=35", "x * 2 = 20"]);
}

#[test]
fn test_py_custom_format_dunder() {
    let src = r#"
class Money:
    def __init__(self, amount, currency="USD"):
        self.amount = amount
        self.currency = currency

    def __format__(self, format_spec):
        if format_spec == "code":
            return f"{self.amount:.2f} {self.currency}"
        elif format_spec == "symbol":
            sym = "$" if self.currency == "USD" else "€"
            return f"{sym}{self.amount:.2f}"
        return f"{self.amount:.2f}"

m = Money(49.5, "USD")
print(f"{m:code}")
print(f"{m:symbol}")
print(f"{m}")
"#;
    assert_eq!(run_python(src), vec!["49.50 USD", "$49.50", "49.50"]);
}

#[test]
fn test_py_fstring_conversion_flags() {
    let src = r#"
text = "hello\nworld"
print(f"{text!r}")
print(f"{text!s}")
print(f"{text!a}")  # ascii
"#;
    assert_eq!(
        run_python(src),
        vec!["'hello\\nworld'", "hello\nworld", "'hello\\nworld'"]
    );
}

#[test]
fn test_py_fstring_datetime_formatting() {
    let src = r#"
from datetime import datetime

dt = datetime(2024, 5, 12, 14, 30, 45)
print(f"{dt:%Y-%m-%d %H:%M:%S}")
print(f"{dt:%B %d, %Y}")
"#;
    assert_eq!(run_python(src), vec!["2024-05-12 14:30:45", "May 12, 2024"]);
}

#[test]
fn test_py_format_method_positional_and_keyword() {
    let src = r#"
print("{0}, {1}, {0}".format("a", "b"))
print("{name}: {score:.1f}".format(name="Alice", score=98.54))
data = {"x": 10, "y": 20}
print("Point({x}, {y})".format(**data))
"#;
    assert_eq!(
        run_python(src),
        vec!["a, b, a", "Alice: 98.5", "Point(10, 20)"]
    );
}

#[test]
fn test_py_string_template_substitutions() {
    let src = r#"
from string import Template

t = Template("$who likes $what")
print(t.substitute(who="Tim", what="apples"))
d = {"who": "Tim"}
print(t.safe_substitute(d))
"#;
    assert_eq!(run_python(src), vec!["Tim likes apples", "Tim likes $what"]);
}
