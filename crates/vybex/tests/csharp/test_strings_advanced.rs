/// C# string operations: interpolation, StringBuilder, methods
/// (Split, Join, Replace, Trim, PadLeft, Format), verbatim strings,
/// char operations, parsing, conversion patterns.

use super::helpers::{compile_csharp_to_wasm, extract_imports, run_csharp};

// ===================================================================
// STRING INTERPOLATION
// ===================================================================

#[test] fn string_interpolation_basic() {
    assert_eq!(run_csharp(r#"
string name = "Alice";
int age = 30;
Console.WriteLine($"{name} is {age}");
"#), &["Alice is 30"]);
}

#[test] fn string_interpolation_expr() {
    assert_eq!(run_csharp(r#"
int x = 5;
Console.WriteLine($"{x} squared is {x * x}");
"#), &["5 squared is 25"]);
}

#[test] fn string_interpolation_method_call() {
    assert_eq!(run_csharp(r#"
string s = "hello";
Console.WriteLine($"upper: {s.ToUpper()}");
"#), &["upper: HELLO"]);
}

#[test] fn string_interpolation_ternary() {
    assert_eq!(run_csharp(r#"
int x = 10;
Console.WriteLine($"x is {(x > 5 ? "big" : "small")}");
"#), &["x is big"]);
}

// ===================================================================
// STRING METHODS
// ===================================================================

#[test] fn string_toupper_tolower() {
    assert_eq!(run_csharp(r#"
Console.WriteLine("Hello World".ToUpper());
Console.WriteLine("Hello World".ToLower());
"#), &["HELLO WORLD", "hello world"]);
}

#[test] fn string_trim() {
    assert_eq!(run_csharp(r#"
string s = "  hello  ";
Console.WriteLine("'" + s.Trim() + "'");
Console.WriteLine("'" + s.TrimStart() + "'");
Console.WriteLine("'" + s.TrimEnd() + "'");
"#), &["'hello'", "'hello  '", "'  hello'"]);
}

#[test] fn string_split() {
    assert_eq!(run_csharp(r#"
string csv = "a,b,c,d";
string[] parts = csv.Split(',');
foreach (var p in parts) Console.WriteLine(p);
"#), &["a", "b", "c", "d"]);
}

#[test] fn string_split_string_separator() {
    assert_eq!(run_csharp(r#"
string s = "one::two::three";
string[] parts = s.Split("::");
foreach (var p in parts) Console.WriteLine(p);
"#), &["one", "two", "three"]);
}

#[test] fn string_join() {
    assert_eq!(run_csharp(r#"
string[] words = { "hello", "world", "test" };
Console.WriteLine(string.Join(", ", words));
"#), &["hello, world, test"]);
}

#[test] fn string_replace() {
    assert_eq!(run_csharp(r#"
string s = "hello world";
Console.WriteLine(s.Replace("world", "there"));
Console.WriteLine(s.Replace("l", "L"));
"#), &["hello there", "heLLo worLd"]);
}

#[test] fn string_contains_startswith_endswith() {
    assert_eq!(run_csharp(r#"
string s = "Hello World";
Console.WriteLine(s.Contains("lo Wo"));
Console.WriteLine(s.StartsWith("Hello"));
Console.WriteLine(s.EndsWith("World"));
Console.WriteLine(s.StartsWith("World"));
"#), &["True", "True", "True", "False"]);
}

#[test] fn string_indexof_lastindexof() {
    assert_eq!(run_csharp(r#"
string s = "abcabc";
Console.WriteLine(s.IndexOf("bc"));
Console.WriteLine(s.LastIndexOf("bc"));
Console.WriteLine(s.IndexOf("xyz"));
"#), &["1", "4", "-1"]);
}

#[test] fn string_substring() {
    assert_eq!(run_csharp(r#"
string s = "Hello World";
Console.WriteLine(s.Substring(6));
Console.WriteLine(s.Substring(0, 5));
"#), &["World", "Hello"]);
}

#[test] fn string_padleft_padright() {
    assert_eq!(run_csharp(r#"
string s = "hi";
Console.WriteLine("'" + s.PadLeft(6) + "'");
Console.WriteLine("'" + s.PadRight(6) + "'");
Console.WriteLine("'" + s.PadLeft(6, '*') + "'");
"#), &["'    hi'", "'hi    '", "'****hi'"]);
}

#[test] fn string_insert_remove() {
    assert_eq!(run_csharp(r#"
string s = "Hello World";
Console.WriteLine(s.Insert(5, " Beautiful"));
Console.WriteLine(s.Remove(5));
Console.WriteLine(s.Remove(5, 1));
"#), &["Hello Beautiful World", "Hello", "HelloWorld"]);
}

#[test] fn string_format() {
    assert_eq!(run_csharp(r#"
Console.WriteLine(string.Format("{0} + {1} = {2}", 1, 2, 3));
Console.WriteLine(string.Format("Name: {0}, Age: {1}", "Bob", 25));
"#), &["1 + 2 = 3", "Name: Bob, Age: 25"]);
}

#[test] fn fully_qualified_system_string_format() {
    assert_eq!(run_csharp(r#"
Console.WriteLine(System.String.Format("{0}-{1}", "A", "B"));
"#), &["A-B"]);
}

#[test] fn string_isnullorempty() {
    assert_eq!(run_csharp(r#"
Console.WriteLine(string.IsNullOrEmpty(""));
Console.WriteLine(string.IsNullOrEmpty(null));
Console.WriteLine(string.IsNullOrEmpty("hello"));
"#), &["true", "true", "false"]);
}

#[test] fn string_isnullorwhitespace() {
    assert_eq!(run_csharp(r#"
Console.WriteLine(string.IsNullOrWhiteSpace("   "));
Console.WriteLine(string.IsNullOrWhiteSpace(""));
Console.WriteLine(string.IsNullOrWhiteSpace("x"));
"#), &["true", "true", "false"]);
}

#[test] fn string_methods_use_stdlib_not_vybe_string_hosts() {
    let wasm = compile_csharp_to_wasm(r#"
string s = "Hello World";
Console.WriteLine(string.IsNullOrEmpty(""));
Console.WriteLine(s.Insert(5, " Beautiful"));
Console.WriteLine(s.Remove(5));
Console.WriteLine(s.Remove(5, 1));
"#);
    let imports = extract_imports(&wasm);
    for forbidden in ["isNullOrEmpty", "insert", "remove"] {
        assert!(
            !imports.iter().any(|(module, name)| module == "vybe:string" && name == forbidden),
            "unexpected vybe:string.{} import in emitted wasm: {:?}",
            forbidden,
            imports
        );
    }
}

// ===================================================================
// STRING BUILDING / CONCATENATION
// ===================================================================

#[test] fn stringbuilder_basic() {
    assert_eq!(run_csharp(r#"
var sb = new System.Text.StringBuilder();
sb.Append("Hello");
sb.Append(" ");
sb.Append("World");
Console.WriteLine(sb.ToString());
"#), &["Hello World"]);
}

#[test] fn stringbuilder_appendline() {
    assert_eq!(run_csharp(r#"
var sb = new System.Text.StringBuilder();
sb.AppendLine("line1");
sb.AppendLine("line2");
Console.Write(sb.ToString());
"#), &["line1", "line2"]);
}

#[test] fn stringbuilder_insert_replace() {
    assert_eq!(run_csharp(r#"
var sb = new System.Text.StringBuilder("Hello World");
sb.Replace("World", "There");
Console.WriteLine(sb.ToString());
sb.Insert(5, " Beautiful");
Console.WriteLine(sb.ToString());
"#), &["Hello There", "Hello Beautiful There"]);
}

#[test] fn string_concat_operator() {
    assert_eq!(run_csharp(r#"
string a = "Hello";
string b = " World";
string c = a + b;
Console.WriteLine(c);
"#), &["Hello World"]);
}

#[test] fn string_concat_method() {
    assert_eq!(run_csharp(r#"
Console.WriteLine(string.Concat("A", "B", "C"));
"#), &["ABC"]);
}

// ===================================================================
// CHAR OPERATIONS
// ===================================================================

#[test] fn char_isupper_islower() {
    assert_eq!(run_csharp(r#"
Console.WriteLine(char.IsUpper('A'));
Console.WriteLine(char.IsLower('a'));
Console.WriteLine(char.IsDigit('5'));
Console.WriteLine(char.IsLetter('x'));
Console.WriteLine(char.IsWhiteSpace(' '));
"#), &["True", "True", "True", "True", "True"]);
}

#[test] fn char_toupper_tolower() {
    assert_eq!(run_csharp(r#"
Console.WriteLine(char.ToUpper('a'));
Console.WriteLine(char.ToLower('Z'));
"#), &["A", "z"]);
}

#[test] fn string_tochararray() {
    assert_eq!(run_csharp(r#"
char[] chars = "hello".ToCharArray();
Array.Reverse(chars);
Console.WriteLine(new string(chars));
"#), &["olleh"]);
}

// ===================================================================
// PARSING AND CONVERSION
// ===================================================================

#[test] fn int_parse_tostring() {
    assert_eq!(run_csharp(r#"
int x = int.Parse("42");
Console.WriteLine(x + 8);
Console.WriteLine(x.ToString());
"#), &["50", "42"]);
}

#[test] fn bool_parse() {
    assert_eq!(run_csharp(r#"
bool t = bool.Parse("True");
bool f = bool.Parse("False");
Console.WriteLine(t);
Console.WriteLine(f);
"#), &["True", "False"]);
}

#[test] fn convert_toint32_tostring() {
    assert_eq!(run_csharp(r#"
int x = Convert.ToInt32("123");
string s = Convert.ToString(456);
Console.WriteLine(x);
Console.WriteLine(s);
"#), &["123", "456"]);
}

// ===================================================================
// VERBATIM AND RAW STRINGS
// ===================================================================

#[test] fn verbatim_string() {
    assert_eq!(run_csharp(r#"
string path = @"C:\Users\test\file.txt";
Console.WriteLine(path);
"#), &["C:\\Users\\test\\file.txt"]);
}

#[test] fn verbatim_string_with_quotes() {
    assert_eq!(run_csharp(r#"
string s = @"He said ""hello""";
Console.WriteLine(s);
"#), &["He said \"hello\""]);
}

#[test] fn escape_sequences() {
    assert_eq!(run_csharp(r#"
Console.WriteLine("tab:\there");
Console.WriteLine("newline done");
"#), &["tab:\there", "newline done"]);
}

// ===================================================================
// STRING COMPARISON
// ===================================================================

#[test] fn string_equals_ignore_case() {
    assert_eq!(run_csharp(r#"
Console.WriteLine(string.Equals("Hello", "hello", StringComparison.OrdinalIgnoreCase));
Console.WriteLine(string.Equals("Hello", "hello"));
"#), &["True", "False"]);
}

#[test] fn string_compareto() {
    assert_eq!(run_csharp(r#"
string a = "apple";
string b = "banana";
Console.WriteLine(a.CompareTo(b) < 0);
Console.WriteLine(b.CompareTo(a) > 0);
Console.WriteLine(a.CompareTo(a) == 0);
"#), &["True", "True", "True"]);
}
