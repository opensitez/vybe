///! Tests for function/sub declarations, returns, recursion, ByRef, and composition.

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
// Return Values
// ============================================================

#[test]
fn test_function_return_explicit() {
    let output = run_vb(r#"
Function Add(a As Integer, b As Integer) As Integer
    Return a + b
End Function

Sub Main()
    Dim result As Integer = Add(3, 4)
    Console.WriteLine(result)
End Sub
"#);
    assert_eq!(output, vec!["7"]);
}

#[test]
fn test_function_return_via_name() {
    let output = run_vb(r#"
Function GetValue() As Integer
    GetValue = 42
End Function

Sub Main()
    Console.WriteLine(GetValue())
End Sub
"#);
    assert_eq!(output, vec!["42"]);
}

#[test]
fn test_recursive_function() {
    let output = run_vb(r#"
Function Factorial(n As Integer) As Integer
    If n <= 1 Then
        Return 1
    Else
        Return n * Factorial(n - 1)
    End If
End Function

Sub Main()
    Console.WriteLine(Factorial(5))
    Console.WriteLine(Factorial(10))
End Sub
"#);
    assert_eq!(output, vec!["120", "3628800"]);
}

// ============================================================
// Sub/Function Composition
// ============================================================

#[test]
fn test_sub_calls_sub() {
    let output = run_vb(r#"
Sub PrintLine(msg As String)
    Console.WriteLine(">> " & msg)
End Sub

Sub Main()
    PrintLine("hello")
    PrintLine("world")
End Sub
"#);
    assert_eq!(output, vec![">> hello", ">> world"]);
}

#[test]
fn test_function_calls_function() {
    let output = run_vb(r#"
Function DoubleIt(x As Integer) As Integer
    Return x * 2
End Function

Function Quadruple(x As Integer) As Integer
    Return DoubleIt(DoubleIt(x))
End Function

Sub Main()
    Console.WriteLine(Quadruple(5))
End Sub
"#);
    assert_eq!(output, vec!["20"]);
}

// ============================================================
// ByRef Parameters
// ============================================================

#[test]
fn test_byref_parameter() {
    let output = run_vb(r#"
Sub Increment(ByRef x As Integer)
    x = x + 1
End Sub

Sub Main()
    Dim n As Integer = 10
    Increment(n)
    Console.WriteLine(n)
    Increment(n)
    Console.WriteLine(n)
End Sub
"#);
    assert_eq!(output, vec!["11", "12"]);
}

// ============================================================
// Exit Sub / Exit Function
// ============================================================

#[test]
fn test_exit_sub() {
    let output = run_vb(r#"
Sub EarlyReturn()
    Console.WriteLine("before")
    Exit Sub
    Console.WriteLine("after")
End Sub

Sub Main()
    EarlyReturn()
    Console.WriteLine("done")
End Sub
"#);
    assert_eq!(output, vec!["before", "done"]);
}

#[test]
fn test_exit_function() {
    let output = run_vb(r#"
Function GetFirst() As String
    Return "first"
    Return "second"
End Function

Sub Main()
    Console.WriteLine(GetFirst())
End Sub
"#);
    assert_eq!(output, vec!["first"]);
}
