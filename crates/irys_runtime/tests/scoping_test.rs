///! Tests for variable scoping rules.

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

#[test]
fn test_variable_scoping_in_for_loop() {
    let output = run_vb(r#"
Sub Main()
    For i As Integer = 1 To 3
        Dim msg As String = "iter " & i
        Console.WriteLine(msg)
    Next
End Sub
"#);
    assert_eq!(output, vec!["iter 1", "iter 2", "iter 3"]);
}

#[test]
fn test_variable_scoping_basic() {
    let output = run_vb(r#"
Sub Main()
    Dim x As Integer = 10
    If True Then
        Dim y As Integer = 20
        Console.WriteLine(x + y)
    End If
    Console.WriteLine(x)
End Sub
"#);
    assert_eq!(output, vec!["30", "10"]);
}

#[test]
fn test_module_level_variable() {
    let output = run_vb(r#"
Module TestModule
    Dim counter As Integer = 0

    Sub IncrementCounter()
        counter = counter + 1
    End Sub

    Sub Main()
        IncrementCounter()
        IncrementCounter()
        IncrementCounter()
        Console.WriteLine(counter)
    End Sub
End Module
"#);
    assert_eq!(output, vec!["3"]);
}
