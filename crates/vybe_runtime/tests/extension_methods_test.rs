use vybe_runtime::{Interpreter, RuntimeSideEffect};
use vybe_parser::ast::Identifier;
use vybe_parser::parse_program;

/// Helper: run a VB source, call Main, and return all console output lines.
fn run_and_capture(source: &str) -> Vec<String> {
    let program = parse_program(source).expect("Failed to parse");
    let mut interp = Interpreter::new();
    interp.run(&program).expect("Runtime error");
    let main_ident = Identifier::new("Main");
    interp.call_procedure(&main_ident, &[]).expect("Failed to call Main");
    interp
        .side_effects
        .iter()
        .filter_map(|e| {
            if let RuntimeSideEffect::ConsoleOutput(msg) = e {
                Some(msg.clone())
            } else {
                None
            }
        })
        .collect()
}

#[test]
fn test_extension_function_on_string() {
    let code = r#"
        <Extension()>
        Function Reverse(s As String) As String
            Dim result As String = ""
            Dim i As Integer
            For i = Len(s) To 1 Step -1
                result = result & Mid(s, i, 1)
            Next
            Return result
        End Function

        Sub Main()
            Dim greeting As String = "Hello"
            Console.WriteLine(greeting.Reverse())
        End Sub
    "#;
    let output = run_and_capture(code);
    assert!(
        output.iter().any(|l| l.contains("olleH")),
        "Expected reversed string 'olleH' in output: {:?}",
        output
    );
}

#[test]
fn test_extension_sub_on_string() {
    let code = r#"
        <Extension()>
        Sub PrintUpper(s As String)
            Console.WriteLine(s.ToUpper())
        End Sub

        Sub Main()
            Dim msg As String = "hello world"
            msg.PrintUpper()
        End Sub
    "#;
    let output = run_and_capture(code);
    assert!(
        output.iter().any(|l| l.contains("HELLO WORLD")),
        "Expected 'HELLO WORLD' in output: {:?}",
        output
    );
}

#[test]
fn test_extension_with_qualified_attribute() {
    let code = r#"
        <Runtime.CompilerServices.Extension()>
        Function IsBlank(s As String) As Boolean
            Return Len(Trim(s)) = 0
        End Function

        Sub Main()
            Dim empty As String = "   "
            Dim full As String = "hi"
            If empty.IsBlank() Then
                Console.WriteLine("BLANK_YES")
            End If
            If Not full.IsBlank() Then
                Console.WriteLine("BLANK_NO")
            End If
        End Sub
    "#;
    let output = run_and_capture(code);
    assert!(
        output.iter().any(|l| l.contains("BLANK_YES")),
        "Expected 'BLANK_YES' in output: {:?}",
        output
    );
    assert!(
        output.iter().any(|l| l.contains("BLANK_NO")),
        "Expected 'BLANK_NO' in output: {:?}",
        output
    );
}

#[test]
fn test_chained_extension_call() {
    let code = r#"
        <Extension()>
        Function Reverse(s As String) As String
            Dim result As String = ""
            Dim i As Integer
            For i = Len(s) To 1 Step -1
                result = result & Mid(s, i, 1)
            Next
            Return result
        End Function

        Sub Main()
            Dim txt As String = "abc"
            ' Reverse twice should yield the original
            Console.WriteLine(txt.Reverse().Reverse())
        End Sub
    "#;
    let output = run_and_capture(code);
    assert!(
        output.iter().any(|l| l.contains("abc")),
        "Expected 'abc' (double reverse) in output: {:?}",
        output
    );
}
