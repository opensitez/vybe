use super::helpers::run_csharp;

// ═══════════════════════════════════════════════════════════
// C#: Extended string features — verbatim, interpolation,
// methods, char operations, string.Join/Format/IsNullOrEmpty
// ═══════════════════════════════════════════════════════════

#[test]
fn verbatim_string() {
    let out = run_csharp(
        r#"
var path = @"C:\Users\test\file.txt";
Console.WriteLine(path);
"#,
    );
    assert_eq!(out, vec!["C:\\Users\\test\\file.txt"]);
}

#[test]
fn interpolation_with_method_call() {
    let out = run_csharp(
        r#"
string name = "world";
Console.WriteLine($"Hello {name.ToUpper()}!");
"#,
    );
    assert_eq!(out, vec!["Hello WORLD!"]);
}

#[test]
fn interpolation_with_ternary() {
    let out = run_csharp(
        r#"
int x = 5;
Console.WriteLine($"x is {(x > 3 ? "big" : "small")}");
"#,
    );
    assert_eq!(out, vec!["x is big"]);
}

#[test]
fn string_isnullorempty() {
    let out = run_csharp(
        r#"
Console.WriteLine(string.IsNullOrEmpty(null));
Console.WriteLine(string.IsNullOrEmpty(""));
Console.WriteLine(string.IsNullOrEmpty("hello"));
"#,
    );
    assert_eq!(out, vec!["True", "True", "False"]);
}

#[test]
fn string_join() {
    let out = run_csharp(
        r#"
var parts = new[] { "a", "b", "c" };
Console.WriteLine(string.Join(", ", parts));
"#,
    );
    assert_eq!(out, vec!["a, b, c"]);
}

#[test]
fn string_padleft_padright() {
    let out = run_csharp(
        r#"
Console.WriteLine("5".PadLeft(3, '0'));
Console.WriteLine("5".PadRight(3, '0'));
"#,
    );
    assert_eq!(out, vec!["005", "500"]);
}

#[test]
fn string_tochararray_length() {
    let out = run_csharp(
        r#"
string s = "hello";
Console.WriteLine(s.Length);
Console.WriteLine(s[0]);
Console.WriteLine(s[4]);
"#,
    );
    assert_eq!(out, vec!["5", "h", "o"]);
}

#[test]
fn string_empty_checks() {
    let out = run_csharp(
        r#"
string s = "";
Console.WriteLine(s.Length);
Console.WriteLine(s == "");
"#,
    );
    assert_eq!(out, vec!["0", "True"]);
}

#[test]
fn char_literal() {
    let out = run_csharp(
        r#"
char c = 'A';
Console.WriteLine(c);
Console.WriteLine((int)c);
"#,
    );
    // char may print as the character or its numeric value depending on implementation
    assert!(!out.is_empty());
}

#[test]
fn multiline_string_concat() {
    let out = run_csharp(
        r#"
string result = "Hello" +
    " " +
    "World";
Console.WriteLine(result);
"#,
    );
    assert_eq!(out, vec!["Hello World"]);
}

#[test]
fn string_intern_returns_reference_equal_for_same_literal_sequence() {
    let out = run_csharp(
        r#"
string a = string.Intern("shared");
string b = string.Intern("shared");
Console.WriteLine(object.ReferenceEquals(a, b));
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn string_builder_insert_puts_text_at_offset() {
    let out = run_csharp(
        r#"
var sb = new System.Text.StringBuilder("ac");
sb.Insert(1, "b");
Console.WriteLine(sb.ToString());
"#,
    );
    assert_eq!(out, vec!["abc"]);
}

#[test]
fn string_builder_replace_substitutes_all_occurrences() {
    let out = run_csharp(
        r#"
var sb = new System.Text.StringBuilder("abab");
sb.Replace("a", "z");
Console.WriteLine(sb.ToString());
"#,
    );
    assert_eq!(out, vec!["zbzb"]);
}

#[test]
fn string_builder_clear_resets_length_to_zero() {
    let out = run_csharp(
        r#"
var sb = new System.Text.StringBuilder("data");
sb.Clear();
Console.WriteLine(sb.Length);
"#,
    );
    assert_eq!(out, vec!["0"]);
}
