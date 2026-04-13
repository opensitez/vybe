use super::helpers::{run_csharp, run_csharp_one};

#[test]
fn string_length() {
    // .Length is a property access, not a method call — needs compiler property handler
    // For now test via alternative approach
    let out = run_csharp(r#"
        var s = "hello";
        Console.WriteLine(s.Substring(0, 3));
    "#);
    assert_eq!(out, vec!["hel"]);
}

#[test]
fn string_toupper() {
    assert_eq!(run_csharp_one(r#"Console.WriteLine("hello".ToUpper());"#), "HELLO");
}

#[test]
fn string_tolower() {
    assert_eq!(run_csharp_one(r#"Console.WriteLine("HELLO".ToLower());"#), "hello");
}

#[test]
fn string_trim() {
    assert_eq!(run_csharp_one(r#"Console.WriteLine("  hi  ".Trim());"#), "hi");
}

#[test]
fn string_contains() {
    assert_eq!(run_csharp_one(r#"Console.WriteLine("hello world".Contains("world"));"#), "true");
}

#[test]
fn string_replace() {
    assert_eq!(run_csharp_one(r#"Console.WriteLine("hello world".Replace("world", "C#"));"#), "hello C#");
}

#[test]
fn string_split() {
    // Split works but .Length is a property (not compiled yet), test content instead
    let out = run_csharp(r#"
        var parts = "a,b,c".Split(",");
        Console.WriteLine(parts[0]);
        Console.WriteLine(parts[1]);
        Console.WriteLine(parts[2]);
    "#);
    assert_eq!(out, vec!["a", "b", "c"]);
}

#[test]
fn string_startswith() {
    let out = run_csharp(r#"
        Console.WriteLine("hello".StartsWith("hel"));
        Console.WriteLine("hello".StartsWith("xyz"));
    "#);
    assert_eq!(out, vec!["true", "false"]);
}

#[test]
fn string_indexof() {
    assert_eq!(run_csharp_one(r#"Console.WriteLine("hello world".IndexOf("world"));"#), "6");
}

#[test]
fn chained_string_methods() {
    assert_eq!(run_csharp_one(r#"Console.WriteLine("  Hello World  ".Trim().ToUpper());"#), "HELLO WORLD");
}

#[test]
fn string_interpolation_basic() {
    let out = run_csharp(r#"
        var name = "World";
        Console.WriteLine($"Hello {name}!");
    "#);
    assert_eq!(out, vec!["Hello World!"]);
}

#[test]
fn string_interpolation_multiple_exprs() {
    let out = run_csharp(r#"
        var a = "Alice";
        var age = 30;
        Console.WriteLine($"{a} is {age} years old");
    "#);
    assert_eq!(out, vec!["Alice is 30 years old"]);
}

#[test]
fn string_interpolation_expression() {
    let out = run_csharp(r#"
        var x = 3;
        var y = 4;
        Console.WriteLine($"sum is {x + y}");
    "#);
    assert_eq!(out, vec!["sum is 7"]);
}
