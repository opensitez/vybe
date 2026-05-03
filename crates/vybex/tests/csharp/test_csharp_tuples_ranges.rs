use super::helpers::run_csharp;

// ═══════════════════════════════════════════════════════════
// C#: Tuples, ranges, type features, conversions
// ═══════════════════════════════════════════════════════════

#[test]
fn tuple_two_elements() {
    let out = run_csharp(r#"
var t = (10, "hello");
Console.WriteLine(t.Item1);
Console.WriteLine(t.Item2);
"#);
    assert_eq!(out, vec!["10", "hello"]);
}

#[test]
fn tuple_three_elements() {
    let out = run_csharp(r#"
var t = (1, 2, 3);
Console.WriteLine(t.Item1 + t.Item2 + t.Item3);
"#);
    assert_eq!(out, vec!["6"]);
}

#[test]
fn int_parse() {
    let out = run_csharp(r#"
int n = int.Parse("42");
Console.WriteLine(n + 8);
"#);
    assert_eq!(out, vec!["50"]);
}

#[test]
fn double_parse() {
    let out = run_csharp(r#"
double d = double.Parse("3.14");
Console.WriteLine(d);
"#);
    assert_eq!(out, vec!["3.14"]);
}

#[test]
fn to_string() {
    let out = run_csharp(r#"
int x = 42;
Console.WriteLine(x.ToString());
Console.WriteLine(42.ToString());
"#);
    assert_eq!(out, vec!["42", "42"]);
}

#[test]
fn range_array_slice() {
    let out = run_csharp(r#"
var arr = new[] { 0, 1, 2, 3, 4 };
var slice = arr[1..4];
foreach (var x in slice) Console.WriteLine(x);
"#);
    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test]
fn numeric_literals() {
    let out = run_csharp(r#"
Console.WriteLine(0xFF);
Console.WriteLine(0b1010);
Console.WriteLine(1.5e2);
"#);
    assert_eq!(out, vec!["255", "10", "150"]);
}

#[test]
fn boolean_literals() {
    let out = run_csharp(r#"
Console.WriteLine(true);
Console.WriteLine(false);
Console.WriteLine(true && false);
Console.WriteLine(true || false);
"#);
    assert_eq!(out, vec!["True", "False", "False", "True"]);
}

#[test]
fn conditional_expression() {
    let out = run_csharp(r#"
int x = 5;
Console.WriteLine(x > 0 ? "positive" : "non-positive");
Console.WriteLine(x > 10 ? "big" : "small");
"#);
    assert_eq!(out, vec!["positive", "small"]);
}

#[test]
fn int_maxvalue() {
    let out = run_csharp(r#"
Console.WriteLine(int.MaxValue);
"#);
    assert_eq!(out, vec!["2147483647"]);
}
