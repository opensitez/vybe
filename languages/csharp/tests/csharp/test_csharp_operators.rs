use super::helpers::run_csharp;

// ═══════════════════════════════════════════════════════════
// C#: Operators — arithmetic, comparison, logical, bitwise,
// null-coalescing, ternary, type checking
// ═══════════════════════════════════════════════════════════

#[test]
fn arithmetic() {
    let out = run_csharp(
        r#"
Console.WriteLine(10 + 5);
Console.WriteLine(10 - 5);
Console.WriteLine(10 * 5);
Console.WriteLine(10 % 3);
"#,
    );
    assert_eq!(out, vec!["15", "5", "50", "1"]);
}

#[test]
fn comparison() {
    let out = run_csharp(
        r#"
Console.WriteLine(1 < 2);
Console.WriteLine(2 > 1);
Console.WriteLine(1 <= 1);
Console.WriteLine(1 >= 1);
Console.WriteLine(1 == 1);
Console.WriteLine(1 != 2);
"#,
    );
    assert_eq!(out, vec!["True", "True", "True", "True", "True", "True"]);
}

#[test]
fn logical_operators() {
    let out = run_csharp(
        r#"
Console.WriteLine(true && true);
Console.WriteLine(true && false);
Console.WriteLine(false || true);
Console.WriteLine(false || false);
Console.WriteLine(!true);
"#,
    );
    assert_eq!(out, vec!["True", "False", "True", "False", "False"]);
}

#[test]
fn null_coalescing() {
    let out = run_csharp(
        r#"
string s = null;
Console.WriteLine(s ?? "default");
s = "hello";
Console.WriteLine(s ?? "default");
"#,
    );
    assert_eq!(out, vec!["default", "hello"]);
}

#[test]
fn compound_assignment() {
    let out = run_csharp(
        r#"
int x = 10;
x += 5; Console.WriteLine(x);
x -= 3; Console.WriteLine(x);
x *= 2; Console.WriteLine(x);
x /= 4; Console.WriteLine(x);
x %= 5; Console.WriteLine(x);
"#,
    );
    assert_eq!(out, vec!["15", "12", "24", "6", "1"]);
}

#[test]
fn increment_decrement() {
    let out = run_csharp(
        r#"
int x = 5;
Console.WriteLine(x++);
Console.WriteLine(x);
Console.WriteLine(++x);
Console.WriteLine(x--);
Console.WriteLine(x);
"#,
    );
    assert_eq!(out, vec!["5", "6", "7", "7", "6"]);
}

#[test]
fn string_concat_operator() {
    let out = run_csharp(
        r#"
Console.WriteLine("Hello" + " " + "World");
Console.WriteLine("num: " + 42);
"#,
    );
    assert_eq!(out, vec!["Hello World", "num: 42"]);
}

#[test]
fn typeof_expression() {
    let out = run_csharp(
        r#"
Console.WriteLine(typeof(int));
Console.WriteLine(typeof(string));
"#,
    );
    assert_eq!(out[0].contains("int") || out[0].contains("Int"), true);
}

#[test]
fn is_type_check() {
    let out = run_csharp(
        r#"
object x = "hello";
Console.WriteLine(x is string);
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn parenthesized_precedence() {
    let out = run_csharp(
        r#"
Console.WriteLine((2 + 3) * 4);
Console.WriteLine(2 + 3 * 4);
"#,
    );
    assert_eq!(out, vec!["20", "14"]);
}

#[test]
fn unary_minus() {
    let out = run_csharp(
        r#"
int x = 42;
Console.WriteLine(-x);
"#,
    );
    assert_eq!(out, vec!["-42"]);
}

#[test]
fn user_defined_plus_operator_combines_struct_fields() {
    let out = run_csharp(
        r#"
struct Vec2 {
    public int X;
    public int Y;
    public static Vec2 operator +(Vec2 a, Vec2 b) =>
        new Vec2 { X = a.X + b.X, Y = a.Y + b.Y };
}
var sum = new Vec2 { X = 1, Y = 2 } + new Vec2 { X = 3, Y = 4 };
Console.WriteLine(sum.X);
Console.WriteLine(sum.Y);
"#,
    );
    assert_eq!(out, vec!["4", "6"]);
}

#[test]
fn user_defined_implicit_conversion_coerces_to_target_type() {
    let out = run_csharp(
        r#"
struct Inch {
    public double Value;
    public static implicit operator double(Inch i) => i.Value;
}
double length = new Inch { Value = 2.5 };
Console.WriteLine(length);
"#,
    );
    assert_eq!(out, vec!["2.5"]);
}

#[test]
fn is_not_pattern_negates_type_test() {
    let out = run_csharp(
        r#"
object value = 7;
Console.WriteLine(value is not string);
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn as_operator_returns_null_for_incompatible_reference_cast() {
    let out = run_csharp(
        r#"
object value = 1;
var text = value as string;
Console.WriteLine(text == null);
"#,
    );
    assert_eq!(out, vec!["True"]);
}
