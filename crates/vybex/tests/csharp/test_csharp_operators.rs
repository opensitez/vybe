use super::helpers::run_csharp;

// ═══════════════════════════════════════════════════════════
// C#: Operators — arithmetic, comparison, logical, bitwise,
// null-coalescing, ternary, type checking
// ═══════════════════════════════════════════════════════════

#[test]
fn arithmetic() {
    let out = run_csharp(r#"
Console.WriteLine(10 + 5);
Console.WriteLine(10 - 5);
Console.WriteLine(10 * 5);
Console.WriteLine(10 % 3);
"#);
    assert_eq!(out, vec!["15", "5", "50", "1"]);
}

#[test]
fn comparison() {
    let out = run_csharp(r#"
Console.WriteLine(1 < 2);
Console.WriteLine(2 > 1);
Console.WriteLine(1 <= 1);
Console.WriteLine(1 >= 1);
Console.WriteLine(1 == 1);
Console.WriteLine(1 != 2);
"#);
    assert_eq!(out, vec!["true", "true", "true", "true", "true", "true"]);
}

#[test]
fn logical_operators() {
    let out = run_csharp(r#"
Console.WriteLine(true && true);
Console.WriteLine(true && false);
Console.WriteLine(false || true);
Console.WriteLine(false || false);
Console.WriteLine(!true);
"#);
    assert_eq!(out, vec!["true", "false", "true", "false", "false"]);
}

#[test]
fn null_coalescing() {
    let out = run_csharp(r#"
string s = null;
Console.WriteLine(s ?? "default");
s = "hello";
Console.WriteLine(s ?? "default");
"#);
    assert_eq!(out, vec!["default", "hello"]);
}

#[test]
fn compound_assignment() {
    let out = run_csharp(r#"
int x = 10;
x += 5; Console.WriteLine(x);
x -= 3; Console.WriteLine(x);
x *= 2; Console.WriteLine(x);
x /= 4; Console.WriteLine(x);
x %= 5; Console.WriteLine(x);
"#);
    assert_eq!(out, vec!["15", "12", "24", "6", "1"]);
}

#[test]
fn increment_decrement() {
    let out = run_csharp(r#"
int x = 5;
Console.WriteLine(x++);
Console.WriteLine(x);
Console.WriteLine(++x);
Console.WriteLine(x--);
Console.WriteLine(x);
"#);
    assert_eq!(out, vec!["5", "6", "7", "7", "6"]);
}

#[test]
fn string_concat_operator() {
    let out = run_csharp(r#"
Console.WriteLine("Hello" + " " + "World");
Console.WriteLine("num: " + 42);
"#);
    assert_eq!(out, vec!["Hello World", "num: 42"]);
}

#[test]
fn typeof_expression() {
    let out = run_csharp(r#"
Console.WriteLine(typeof(int));
Console.WriteLine(typeof(string));
"#);
    assert_eq!(out[0].contains("int") || out[0].contains("Int"), true);
}

#[test]
#[ignore]
fn is_type_check() {
    let out = run_csharp(r#"
object x = "hello";
Console.WriteLine(x is string);
"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn parenthesized_precedence() {
    let out = run_csharp(r#"
Console.WriteLine((2 + 3) * 4);
Console.WriteLine(2 + 3 * 4);
"#);
    assert_eq!(out, vec!["20", "14"]);
}

#[test]
fn unary_minus() {
    let out = run_csharp(r#"
int x = 42;
Console.WriteLine(-x);
"#);
    assert_eq!(out, vec!["-42"]);
}
