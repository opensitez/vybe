use super::helpers::run_csharp;

// ═══════════════════════════════════════════════════════════
// C#: Strings — methods, interpolation, formatting
// ═══════════════════════════════════════════════════════════

#[test]
fn string_length() {
    let out = run_csharp(r#"
Console.WriteLine("hello".Length);
"#);
    assert_eq!(out, vec!["5"]);
}

#[test]
fn string_toupper_tolower() {
    let out = run_csharp(r#"
Console.WriteLine("hello".ToUpper());
Console.WriteLine("HELLO".ToLower());
"#);
    assert_eq!(out, vec!["HELLO", "hello"]);
}

#[test]
fn string_trim() {
    let out = run_csharp(r#"
Console.WriteLine("  hello  ".Trim());
"#);
    assert_eq!(out, vec!["hello"]);
}

#[test]
fn string_contains() {
    let out = run_csharp(r#"
Console.WriteLine("hello world".Contains("world"));
Console.WriteLine("hello world".Contains("xyz"));
"#);
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn string_replace() {
    let out = run_csharp(r#"
Console.WriteLine("hello world".Replace("world", "C#"));
"#);
    assert_eq!(out, vec!["hello C#"]);
}

#[test]
fn string_split() {
    let out = run_csharp(r#"
string[] parts = "a,b,c".Split(",");
Console.WriteLine(parts.Length);
Console.WriteLine(parts[1]);
"#);
    assert_eq!(out, vec!["3", "b"]);
}

#[test]
fn string_startswith_endswith() {
    let out = run_csharp(r#"
Console.WriteLine("hello".StartsWith("hel"));
Console.WriteLine("hello".EndsWith("llo"));
Console.WriteLine("hello".StartsWith("xyz"));
"#);
    assert_eq!(out, vec!["True", "True", "False"]);
}

#[test]
fn string_indexof() {
    let out = run_csharp(r#"
Console.WriteLine("hello world".IndexOf("world"));
Console.WriteLine("hello world".IndexOf("xyz"));
"#);
    assert_eq!(out, vec!["6", "-1"]);
}

#[test]
fn string_substring() {
    let out = run_csharp(r#"
Console.WriteLine("hello world".Substring(6));
Console.WriteLine("hello world".Substring(0, 5));
"#);
    assert_eq!(out, vec!["world", "hello"]);
}

#[test]
fn string_interpolation() {
    let out = run_csharp(r#"
string name = "Alice";
int age = 30;
Console.WriteLine($"{name} is {age}");
"#);
    assert_eq!(out, vec!["Alice is 30"]);
}

#[test]
fn string_interpolation_expression() {
    let out = run_csharp(r#"
int a = 3, b = 4;
Console.WriteLine($"sum = {a + b}");
"#);
    assert_eq!(out, vec!["sum = 7"]);
}

#[test]
fn string_concat_plus() {
    let out = run_csharp(r#"
Console.WriteLine("Hello" + " " + "World");
"#);
    assert_eq!(out, vec!["Hello World"]);
}

#[test]
fn string_comparison() {
    let out = run_csharp(r#"
Console.WriteLine("abc" == "abc");
Console.WriteLine("abc" != "xyz");
"#);
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn string_chained_methods() {
    let out = run_csharp(r#"
Console.WriteLine("  Hello World  ".Trim().ToUpper());
"#);
    assert_eq!(out, vec!["HELLO WORLD"]);
}
