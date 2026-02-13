///! Tests for built-in Math functions.

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
fn test_math_abs() {
    let output = run_vb(r#"
Sub Main()
    Console.WriteLine(Math.Abs(-5))
    Console.WriteLine(Math.Abs(5))
    Console.WriteLine(Math.Abs(0))
End Sub
"#);
    assert_eq!(output, vec!["5", "5", "0"]);
}

#[test]
fn test_math_max_min() {
    let output = run_vb(r#"
Sub Main()
    Console.WriteLine(Math.Max(3, 7))
    Console.WriteLine(Math.Min(3, 7))
End Sub
"#);
    assert_eq!(output, vec!["7", "3"]);
}
