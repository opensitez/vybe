use vybe_runtime::{Interpreter, RuntimeSideEffect, Value};
use vybe_parser::parse_program;
use vybe_parser::ast::Identifier;

fn run_vb(code: &str) -> (Interpreter, Vec<String>) {
    let program = parse_program(code).expect("Parse error");
    let mut interp = Interpreter::new();
    interp.run(&program).expect("Runtime error");
    interp.call_procedure(&Identifier::new("Main"), &[]).expect("Failed to call Main");
    let output = interp.side_effects.iter().filter_map(|e| {
        if let RuntimeSideEffect::ConsoleOutput(msg) = e { Some(msg.trim_end().to_string()) } else { None }
    }).collect();
    (interp, output)
}

#[test]
fn test_add_handler_runtime() {
    let code = r#"
Public Class Form1
    Public Sub Register()
        AddHandler Button1.Click, AddressOf HandleClick
    End Sub

    Public Sub HandleClick()
        Console.WriteLine("Clicked!")
    End Sub
End Class

Sub Main()
    Dim f As New Form1()
    f.Register()
End Sub
"#;
    let (interp, _) = run_vb(code);
    
    // Verify handler is registered in EventSystem
    let handlers = interp.get_event_handlers("form1", "button1", "click");
    assert_eq!(handlers.len(), 1);
    assert!(handlers[0].eq_ignore_ascii_case("HandleClick"));
}

#[test]
fn test_application_run_side_effect() {
    let code = r#"
Public Class Form1
End Class

Sub Main()
    Dim f As New Form1()
    Application.Run(f)
End Sub
"#;
    let (interp, _) = run_vb(code);
    
    // Check for RunApplication side effect
    let has_run = interp.side_effects.iter().any(|e| {
        if let RuntimeSideEffect::RunApplication { form_name } = e {
            form_name == "Form1"
        } else {
            false
        }
    });
    assert!(has_run, "RunApplication side effect not found");
}

#[test]
fn test_application_exit_side_effect() {
    let code = r#"
Public Class Form1
End Class

Sub Main()
    Try
        ' We need to set __form_instance__ for Exit to work (simulating running app)
        Application.Run(New Form1()) 
        Application.Exit()
    Catch e As Exception
        Console.WriteLine(e.Message)
    End Try
End Sub
"#;
    let (interp, _) = run_vb(code);
    
    // Check for FormClose side effect
    let has_close = interp.side_effects.iter().any(|e| {
        if let RuntimeSideEffect::FormClose { form_name } = e {
            form_name == "Form1"
        } else {
            false
        }
    });
    assert!(has_close, "FormClose side effect not found");
}
