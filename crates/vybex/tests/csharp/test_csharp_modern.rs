use super::helpers::run_csharp;

// ═══════════════════════════════════════════════════════════
// C#: Modern features — records, tuples, var, type features,
// null handling, generics
// ═══════════════════════════════════════════════════════════

#[test]
fn var_declaration() {
    let out = run_csharp(r#"
var x = 42;
var s = "hello";
Console.WriteLine(x);
Console.WriteLine(s);
"#);
    assert_eq!(out, vec!["42", "hello"]);
}

#[test]
fn record_basic() {
    let out = run_csharp(r#"
record Point(int X, int Y);
var p = new Point(3, 4);
Console.WriteLine(p.X);
Console.WriteLine(p.Y);
"#);
    assert_eq!(out, vec!["3", "4"]);
}

#[test]
fn multiple_return_paths() {
    let out = run_csharp(r#"
string Classify(int x) {
    if (x > 0) return "positive";
    if (x < 0) return "negative";
    return "zero";
}
Console.WriteLine(Classify(5));
Console.WriteLine(Classify(-3));
Console.WriteLine(Classify(0));
"#);
    assert_eq!(out, vec!["positive", "negative", "zero"]);
}

#[test]
fn default_parameters() {
    let out = run_csharp(r#"
string Greet(string name = "World") {
    return "Hello " + name;
}
Console.WriteLine(Greet());
Console.WriteLine(Greet("Alice"));
"#);
    assert_eq!(out, vec!["Hello World", "Hello Alice"]);
}

#[test]
fn params_array() {
    let out = run_csharp(r#"
int Sum(params int[] nums) {
    int total = 0;
    foreach (var n in nums) total += n;
    return total;
}
Console.WriteLine(Sum(1, 2, 3, 4));
"#);
    assert_eq!(out, vec!["10"]);
}

#[test]
fn null_conditional() {
    let out = run_csharp(r#"
string s = null;
Console.WriteLine(s?.Length);
s = "hello";
Console.WriteLine(s?.Length);
"#);
    // null?.Length should be null/nothing
    assert_eq!(out.len(), 2);
}

#[test]
fn boolean_values() {
    let out = run_csharp(r#"
bool t = true;
bool f = false;
Console.WriteLine(t);
Console.WriteLine(f);
Console.WriteLine(t && f);
Console.WriteLine(t || f);
Console.WriteLine(!t);
"#);
    assert_eq!(out, vec!["True", "False", "False", "True", "False"]);
}

#[test]
fn math_operations() {
    let out = run_csharp(r#"
Console.WriteLine(Math.Abs(-42));
Console.WriteLine(Math.Max(10, 20));
Console.WriteLine(Math.Min(10, 20));
Console.WriteLine(Math.Sqrt(25));
Console.WriteLine(Math.Floor(3.7));
Console.WriteLine(Math.Ceiling(3.2));
"#);
    assert_eq!(out, vec!["42", "20", "10", "5", "3", "4"]);
}

#[test]
fn object_initializer() {
    let out = run_csharp(r#"
class Point {
    public int X { get; set; }
    public int Y { get; set; }
}
var p = new Point { X = 10, Y = 20 };
Console.WriteLine(p.X + p.Y);
"#);
    assert_eq!(out, vec!["30"]);
}

#[test]
fn using_static_class() {
    let out = run_csharp(r#"
static class StringUtils {
    public static string Reverse(string s) {
        string result = "";
        for (int i = s.Length - 1; i >= 0; i--) {
            result += s[i];
        }
        return result;
    }
}
Console.WriteLine(StringUtils.Reverse("hello"));
"#);
    assert_eq!(out, vec!["olleh"]);
}
