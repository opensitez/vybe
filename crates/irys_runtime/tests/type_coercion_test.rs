///! Tests for type coercion, conversions, enums, Nothing, and boolean operations.

use irys_runtime::{Interpreter, RuntimeSideEffect};
use irys_parser::ast::Identifier;
use irys_parser::parse_program;

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
// Implicit Coercion
// ============================================================

#[test]
fn test_integer_to_double_coercion() {
    let output = run_vb(r#"
Sub Main()
    Dim x As Double = 5
    Console.WriteLine(x)
    Dim y As Double = 3 + 0.14
    Console.WriteLine(y)
End Sub
"#);
    assert_eq!(output[0], "5");
    assert!(output[1].starts_with("3.14"));
}

#[test]
fn test_string_to_number_coercion() {
    let output = run_vb(r#"
Sub Main()
    Dim x As Integer = CInt("42")
    Console.WriteLine(x)
    Dim y As Double = CDbl("3.14")
    Console.WriteLine(y)
End Sub
"#);
    assert_eq!(output[0], "42");
    assert!(output[1].starts_with("3.14"));
}

#[test]
fn test_boolean_to_string() {
    let output = run_vb(r#"
Sub Main()
    Console.WriteLine(CStr(True))
    Console.WriteLine(CStr(False))
End Sub
"#);
    assert_eq!(output, vec!["True", "False"]);
}

// ============================================================
// Explicit Conversion Functions
// ============================================================

#[test]
fn test_cint_cdbl_cstr_cbool() {
    let output = run_vb(r#"
Sub Main()
    Console.WriteLine(CInt(42.0))
    Console.WriteLine(CDbl(42))
    Console.WriteLine(CStr(123))
    Console.WriteLine(CBool(1))
    Console.WriteLine(CBool(0))
End Sub
"#);
    assert_eq!(output[0], "42");
    assert_eq!(output[1], "42");
    assert_eq!(output[2], "123");
    assert_eq!(output[3], "True");
    assert_eq!(output[4], "False");
}

// ============================================================
// Enumerations
// ============================================================

#[test]
fn test_enum_basic() {
    let output = run_vb(r#"
Enum Color
    Red = 1
    Green = 2
    Blue = 3
End Enum

Sub Main()
    Dim c As Color = Color.Green
    Console.WriteLine(c)
    If c = Color.Green Then
        Console.WriteLine("is green")
    End If
End Sub
"#);
    assert!(output.iter().any(|s| s.contains("2") || s.contains("Green")));
    assert!(output.iter().any(|s| s == "is green"));
}

// ============================================================
// Nothing / Null Handling
// ============================================================

#[test]
fn test_nothing_assignment() {
    let output = run_vb(r#"
Sub Main()
    Dim obj As Object = Nothing
    If obj Is Nothing Then
        Console.WriteLine("is nothing")
    End If
End Sub
"#);
    assert_eq!(output, vec!["is nothing"]);
}
