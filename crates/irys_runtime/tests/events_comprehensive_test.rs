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
// AddHandler / Event Handler calling
// ============================================================

#[test]
fn test_class_with_handles() {
    // Test that Sub with Handles clause is properly parsed  
    let code = r#"
Public Class MyForm
    Public Sub Button1_Click() Handles Button1.Click
        Console.WriteLine("button clicked")
    End Sub
End Class

Sub Main()
    Console.WriteLine("parsed ok")
End Sub
"#;
    let output = run_vb(code);
    assert_eq!(output, vec!["parsed ok"]);
}

#[test]
fn test_addhandler_statement_parses() {
    let code = r#"
Public Class MyForm
    Public Sub DoInit()
        AddHandler Button1.Click, AddressOf Button1_Click
    End Sub

    Public Sub Button1_Click()
    End Sub
End Class

Sub Main()
    Console.WriteLine("addhandler parsed")
End Sub
"#;
    let output = run_vb(code);
    assert_eq!(output, vec!["addhandler parsed"]);
}

#[test]
fn test_removehandler_statement_parses() {
    let code = r#"
Public Class MyForm
    Public Sub Cleanup()
        RemoveHandler Button1.Click, AddressOf Button1_Click
    End Sub

    Public Sub Button1_Click()
    End Sub
End Class

Sub Main()
    Console.WriteLine("removehandler parsed")
End Sub
"#;
    let output = run_vb(code);
    assert_eq!(output, vec!["removehandler parsed"]);
}

// ============================================================
// RaiseEvent Parsing (execution in module context)
// ============================================================

#[test]
fn test_raiseevent_parses_in_class() {
    let code = r#"
Public Class MyClass
    Public Event ValueChanged As EventHandler

    Public Sub SetValue(v As Integer)
        RaiseEvent ValueChanged()
    End Sub
End Class

Sub Main()
    Console.WriteLine("raiseevent ok")
End Sub
"#;
    let output = run_vb(code);
    assert_eq!(output, vec!["raiseevent ok"]);
}

// ============================================================
// Event Handler via Sub Handles (common WinForms pattern)
// ============================================================

#[test]
fn test_sub_with_multiple_handles() {
    let code = r#"
Public Class MyForm
    Public Sub AllButtons_Click() Handles Button1.Click, Button2.Click
    End Sub
End Class

Sub Main()
    Console.WriteLine("multi-handles ok")
End Sub
"#;
    let output = run_vb(code);
    assert_eq!(output, vec!["multi-handles ok"]);
}

// ============================================================
// Event Declaration parsing
// ============================================================

#[test]
fn test_event_declaration_in_class() {
    let code = r#"
Public Class Observable
    Public Event PropertyChanged As EventHandler
    Public Event DataReady As EventHandler
End Class

Sub Main()
    Console.WriteLine("events declared")
End Sub
"#;
    let output = run_vb(code);
    assert_eq!(output, vec!["events declared"]);
}

// ============================================================
// Event calling through class methods
// ============================================================

#[test]
fn test_class_method_call_with_event_sub() {
    let code = r#"
Public Class Counter
    Public Value As Integer = 0

    Public Sub Increment()
        Value = Value + 1
    End Sub
End Class

Sub Main()
    Dim c As New Counter()
    c.Increment()
    c.Increment()
    c.Increment()
    Console.WriteLine(c.Value)
End Sub
"#;
    let output = run_vb(code);
    assert_eq!(output, vec!["3"]);
}
