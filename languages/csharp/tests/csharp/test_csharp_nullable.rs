use super::helpers::run_csharp;

// ═══════════════════════════════════════════════════════════
// C#: Nullable types, null handling, null-coalescing,
// null-conditional, patterns
// ═══════════════════════════════════════════════════════════

#[test]
fn null_coalescing_string() {
    let out = run_csharp(
        r#"
string s = null;
Console.WriteLine(s ?? "fallback");
s = "hello";
Console.WriteLine(s ?? "fallback");
"#,
    );
    assert_eq!(out, vec!["fallback", "hello"]);
}

#[test]
fn null_coalescing_assign() {
    let out = run_csharp(
        r#"
string s = null;
s ??= "assigned";
Console.WriteLine(s);
s ??= "not this";
Console.WriteLine(s);
"#,
    );
    assert_eq!(out, vec!["assigned", "assigned"]);
}

#[test]
fn null_conditional_member() {
    let out = run_csharp(
        r#"
class Wrapper {
    public string Value;
    public Wrapper(string v) { Value = v; }
}
Wrapper w = null;
Console.WriteLine(w?.Value ?? "null");
w = new Wrapper("hello");
Console.WriteLine(w?.Value ?? "null");
"#,
    );
    assert_eq!(out, vec!["null", "hello"]);
}

#[test]
fn null_check_with_if() {
    let out = run_csharp(
        r#"
string s = null;
if (s == null) {
    Console.WriteLine("is null");
} else {
    Console.WriteLine("has value");
}
s = "test";
if (s != null) {
    Console.WriteLine("has value");
}
"#,
    );
    assert_eq!(out, vec!["is null", "has value"]);
}

#[test]
fn null_in_ternary() {
    let out = run_csharp(
        r#"
string s = null;
Console.WriteLine(s != null ? s : "none");
s = "found";
Console.WriteLine(s != null ? s : "none");
"#,
    );
    assert_eq!(out, vec!["none", "found"]);
}

#[test]
fn null_object_pattern() {
    let out = run_csharp(
        r#"
class Box {
    public int Value;
    public Box(int v) { Value = v; }
}
Box b = null;
Console.WriteLine(b == null);
b = new Box(42);
Console.WriteLine(b == null);
Console.WriteLine(b.Value);
"#,
    );
    assert_eq!(out, vec!["True", "False", "42"]);
}
