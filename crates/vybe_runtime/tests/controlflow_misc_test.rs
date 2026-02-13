///! Tests for miscellaneous control flow: constants, If expressions, nested loops.

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

#[test]
fn test_constants() {
    let output = run_vb(r#"
Sub Main()
    Const PI As Double = 3.14159
    Const MAX_SIZE As Integer = 100
    Console.WriteLine(PI)
    Console.WriteLine(MAX_SIZE)
End Sub
"#);
    assert!(output[0].starts_with("3.14159"));
    assert_eq!(output[1], "100");
}

#[test]
fn test_if_expression() {
    let output = run_vb(r#"
Sub Main()
    Dim x As Integer = 5
    Dim result As String = If(x > 3, "big", "small")
    Console.WriteLine(result)
    result = If(x > 10, "big", "small")
    Console.WriteLine(result)
End Sub
"#);
    assert_eq!(output, vec!["big", "small"]);
}

#[test]
fn test_nested_for_loops() {
    let output = run_vb(r#"
Sub Main()
    Dim result As String = ""
    For i As Integer = 1 To 3
        For j As Integer = 1 To 3
            result &= "(" & i & "," & j & ") "
        Next
    Next
    Console.WriteLine(result)
End Sub
"#);
    assert_eq!(output, vec!["(1,1) (1,2) (1,3) (2,1) (2,2) (2,3) (3,1) (3,2) (3,3)"]);
}
