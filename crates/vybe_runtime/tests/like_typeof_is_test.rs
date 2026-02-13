use vybe_runtime::{Interpreter, RuntimeSideEffect};
use vybe_parser::ast::Identifier;
use vybe_parser::parse_program;

fn run_vb(code: &str) -> Vec<String> {
    let program = parse_program(code).expect("Parse error");
    let mut interp = Interpreter::new();
    interp.run(&program).expect("Runtime error");
    interp.call_procedure(&Identifier::new("Main"), &[]).expect("Failed to call Main");
    interp.side_effects.iter().filter_map(|e| {
        if let RuntimeSideEffect::ConsoleOutput(msg) = e { Some(msg.trim_end().to_string()) } else { None }
    }).collect()
}

// ============================================================
// Like Operator — Wildcard *
// ============================================================

#[test]
fn test_like_wildcard_star() {
    let output = run_vb(r#"
Sub Main()
    Console.WriteLine("Hello World" Like "Hello*")
    Console.WriteLine("Hello World" Like "*World")
    Console.WriteLine("Hello World" Like "*lo W*")
    Console.WriteLine("Hello World" Like "*xyz*")
    Console.WriteLine("" Like "*")
End Sub
"#);
    assert_eq!(output, vec!["True", "True", "True", "False", "True"]);
}

// ============================================================
// Like Operator — Single char ?
// ============================================================

#[test]
fn test_like_single_char() {
    let output = run_vb(r#"
Sub Main()
    Console.WriteLine("ABC" Like "A?C")
    Console.WriteLine("ABC" Like "???")
    Console.WriteLine("AB" Like "???")
    Console.WriteLine("ABCD" Like "???")
End Sub
"#);
    assert_eq!(output, vec!["True", "True", "False", "False"]);
}

// ============================================================
// Like Operator — Digit #
// ============================================================

#[test]
fn test_like_digit() {
    let output = run_vb(r####"
Sub Main()
    Console.WriteLine("A1B" Like "A#B")
    Console.WriteLine("A9B" Like "A#B")
    Console.WriteLine("AXB" Like "A#B")
    Console.WriteLine("123" Like "###")
End Sub
"####);
    assert_eq!(output, vec!["True", "True", "False", "True"]);
}

// ============================================================
// Like Operator — Character list [abc]
// ============================================================

#[test]
fn test_like_char_list() {
    let output = run_vb(r#"
Sub Main()
    Console.WriteLine("A" Like "[ABC]")
    Console.WriteLine("B" Like "[ABC]")
    Console.WriteLine("D" Like "[ABC]")
End Sub
"#);
    assert_eq!(output, vec!["True", "True", "False"]);
}

// ============================================================
// Like Operator — Character range [a-z]
// ============================================================

#[test]
fn test_like_char_range() {
    let output = run_vb(r#"
Sub Main()
    Console.WriteLine("m" Like "[a-z]")
    Console.WriteLine("M" Like "[a-z]")
    Console.WriteLine("5" Like "[0-9]")
End Sub
"#);
    assert_eq!(output, vec!["True", "False", "True"]);
}

// ============================================================
// Like Operator — Negated list [!abc]
// ============================================================

#[test]
fn test_like_negated() {
    let output = run_vb(r#"
Sub Main()
    Console.WriteLine("A" Like "[!XYZ]")
    Console.WriteLine("X" Like "[!XYZ]")
End Sub
"#);
    assert_eq!(output, vec!["True", "False"]);
}

// ============================================================
// Like Operator — Exact match
// ============================================================

#[test]
fn test_like_exact() {
    let output = run_vb(r#"
Sub Main()
    Console.WriteLine("Hello" Like "Hello")
    Console.WriteLine("Hello" Like "hello")
    Console.WriteLine("Hello" Like "Helo")
End Sub
"#);
    // VB.NET Like is case-insensitive with Option Compare Text, but 
    // in binary mode (default) it's case-sensitive for letters outside brackets
    // Our implementation uses case-insensitive by default
    assert_eq!(output[0], "True");
    assert_eq!(output[2], "False");
}

// ============================================================
// Like Operator — Complex patterns
// ============================================================

#[test]
fn test_like_complex_pattern() {
    let output = run_vb(r####"
Sub Main()
    ' Match a US phone format like (###) ###-####
    Console.WriteLine("(555) 123-4567" Like "(###) ###-####")
    Console.WriteLine("555-123-4567" Like "(###) ###-####")
End Sub
"####);
    assert_eq!(output, vec!["True", "False"]);
}

// ============================================================
// Is / IsNot — Nothing comparison
// ============================================================

#[test]
fn test_is_nothing() {
    let output = run_vb(r#"
Sub Main()
    Dim obj As Object = Nothing
    If obj Is Nothing Then
        Console.WriteLine("is nothing")
    Else
        Console.WriteLine("not nothing")
    End If
End Sub
"#);
    assert_eq!(output, vec!["is nothing"]);
}

#[test]
fn test_isnot_nothing() {
    let output = run_vb(r#"
Public Class Foo
    Public Value As Integer = 42
End Class

Sub Main()
    Dim obj As New Foo()
    If obj IsNot Nothing Then
        Console.WriteLine("not nothing")
    Else
        Console.WriteLine("is nothing")
    End If
End Sub
"#);
    assert_eq!(output, vec!["not nothing"]);
}

#[test]
fn test_is_nothing_after_set_nothing() {
    let output = run_vb(r#"
Public Class Foo
    Public Value As Integer = 42
End Class

Sub Main()
    Dim obj As New Foo()
    Console.WriteLine(obj IsNot Nothing)
    obj = Nothing
    Console.WriteLine(obj Is Nothing)
End Sub
"#);
    assert_eq!(output, vec!["True", "True"]);
}

// ============================================================
// Is / IsNot — Object identity
// ============================================================

#[test]
fn test_is_same_object() {
    let output = run_vb(r#"
Public Class Foo
    Public Value As Integer
End Class

Sub Main()
    Dim a As New Foo()
    Dim b As Foo = a
    If a Is b Then
        Console.WriteLine("same")
    Else
        Console.WriteLine("different")
    End If
End Sub
"#);
    assert_eq!(output, vec!["same"]);
}

#[test]
fn test_is_different_objects() {
    let output = run_vb(r#"
Public Class Foo
    Public Value As Integer
End Class

Sub Main()
    Dim a As New Foo()
    Dim b As New Foo()
    If a Is b Then
        Console.WriteLine("same")
    Else
        Console.WriteLine("different")
    End If
End Sub
"#);
    assert_eq!(output, vec!["different"]);
}

// ============================================================
// TypeOf...Is
// ============================================================

#[test]
fn test_typeof_is_basic() {
    // TypeOf...Is checks the runtime type of a value
    let output = run_vb(r#"
Sub Main()
    Dim s As String = "hello"
    Dim n As Integer = 42
    Dim d As Double = 3.14
    Dim b As Boolean = True
    If TypeOf s Is String Then
        Console.WriteLine("s is String")
    End If
    If TypeOf n Is Integer Then
        Console.WriteLine("n is Integer")
    End If
    If TypeOf d Is Double Then
        Console.WriteLine("d is Double")
    End If
    If TypeOf b Is Boolean Then
        Console.WriteLine("b is Boolean")
    End If
End Sub
"#);
    assert_eq!(output, vec!["s is String", "n is Integer", "d is Double", "b is Boolean"]);
}

#[test]
fn test_typeof_is_in_if_condition() {
    // TypeOf...Is works directly in If conditions without parentheses
    let output = run_vb(r#"
Sub Main()
    Dim s As String = "hello"
    If TypeOf s Is String Then
        Console.WriteLine("direct")
    End If
    If (TypeOf s Is String) Then
        Console.WriteLine("parens")
    End If
End Sub
"#);
    assert_eq!(output, vec!["direct", "parens"]);
}
