///! Tests for string operations and string interpolation.

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
fn test_string_methods() {
    let output = run_vb(r#"
Sub Main()
    Dim s As String = "Hello World"
    Console.WriteLine(s.Length)
    Console.WriteLine(s.ToUpper())
    Console.WriteLine(s.ToLower())
    Console.WriteLine(s.Substring(0, 5))
    Console.WriteLine(s.Contains("World"))
    Console.WriteLine(s.IndexOf("World"))
    Console.WriteLine(s.Replace("World", "VB.NET"))
End Sub
"#);
    assert_eq!(output[0], "11");
    assert_eq!(output[1], "HELLO WORLD");
    assert_eq!(output[2], "hello world");
    assert_eq!(output[3], "Hello");
    assert_eq!(output[4], "True");
    assert_eq!(output[5], "6");
    assert_eq!(output[6], "Hello VB.NET");
}

#[test]
fn test_string_interpolation() {
    let output = run_vb(r#"
Sub Main()
    Dim name As String = "World"
    Dim age As Integer = 42
    Console.WriteLine($"Hello {name}!")
    Console.WriteLine($"Age is {age}")
    Console.WriteLine($"Sum: {1 + 2}")
End Sub
"#);
    assert_eq!(output, vec!["Hello World!", "Age is 42", "Sum: 3"]);
}
