use irys_runtime::{Interpreter, RuntimeSideEffect};
use irys_parser::ast::Identifier;
use irys_parser::parse_program;

fn run_vb_and_collect(code: &str) -> Vec<String> {
    let program = parse_program(code).expect("Parse error");
    let mut interp = Interpreter::new();
    interp.run(&program).expect("Runtime error");
    interp.call_procedure(&Identifier::new("Main"), &[]).expect("Failed to call Main");
    interp.side_effects.iter().filter_map(|e| {
        if let RuntimeSideEffect::ConsoleOutput(msg) = e { Some(msg.trim_end().to_string()) } else { None }
    }).collect()
}

fn run_vb_file_and_collect(path: &str) -> Vec<String> {
    let source = std::fs::read_to_string(path).expect(&format!("Failed to read {}", path));
    run_vb_and_collect(&source)
}

fn assert_all_pass(output: &[String]) {
    let pass_count = output.iter().filter(|s| s.starts_with("PASS:")).count();
    let fail_lines: Vec<&String> = output.iter().filter(|s| s.starts_with("FAIL:")).collect();
    assert!(fail_lines.is_empty(), "Failures found:\n{}", fail_lines.iter().map(|s| s.as_str()).collect::<Vec<_>>().join("\n"));
    assert!(pass_count > 0, "No PASS assertions found in output");
}

// ============================================================
// Operator Precedence Tests (from .vb file)
// ============================================================

#[test]
fn test_operator_precedence_file() {
    let output = run_vb_file_and_collect("../../tests/test_operator_precedence.vb");
    assert!(output.iter().any(|s| s.contains("=== Operator Precedence Tests ===")));
    assert!(output.iter().any(|s| s.contains("=== Operator Precedence Tests Done ===")));
    assert_all_pass(&output);
}

// ============================================================
// Integer Division
// ============================================================

#[test]
fn test_integer_division_basic() {
    let output = run_vb_and_collect(r#"
Sub Main()
    Console.WriteLine(10 \ 3)
    Console.WriteLine(7 \ 2)
    Console.WriteLine(100 \ 10)
    Console.WriteLine(1 \ 2)
End Sub
"#);
    assert_eq!(output, vec!["3", "3", "10", "0"]);
}

#[test]
fn test_integer_division_vs_regular_division() {
    let output = run_vb_and_collect(r#"
Sub Main()
    Dim intResult As Integer = 7 \ 2
    Dim floatResult As Double = 7 / 2
    Console.WriteLine("Int: " & intResult)
    Console.WriteLine("Float: " & floatResult)
End Sub
"#);
    assert!(output[0].contains("Int: 3"));
    assert!(output[1].contains("Float: 3.5"));
}

// ============================================================
// Exponent
// ============================================================

#[test]
fn test_exponent_operator() {
    let output = run_vb_and_collect(r#"
Sub Main()
    Console.WriteLine(2 ^ 0)
    Console.WriteLine(2 ^ 1)
    Console.WriteLine(2 ^ 10)
    Console.WriteLine(3 ^ 3)
End Sub
"#);
    assert_eq!(output, vec!["1", "2", "1024", "27"]);
}

#[test]
fn test_exponent_precedence_higher_than_multiply() {
    let output = run_vb_and_collect(r#"
Sub Main()
    ' 4 * 2 ^ 3 should be 4 * 8 = 32
    Console.WriteLine(4 * 2 ^ 3)
    ' 2 ^ 3 + 1 should be 8 + 1 = 9
    Console.WriteLine(2 ^ 3 + 1)
End Sub
"#);
    assert_eq!(output, vec!["32", "9"]);
}

// ============================================================
// AndAlso / OrElse (short-circuit)
// ============================================================

#[test]
fn test_andalso_short_circuit() {
    let output = run_vb_and_collect(r#"
Sub Main()
    Console.WriteLine(True AndAlso True)
    Console.WriteLine(True AndAlso False)
    Console.WriteLine(False AndAlso True)
    Console.WriteLine(False AndAlso False)
End Sub
"#);
    assert_eq!(output, vec!["True", "False", "False", "False"]);
}

#[test]
fn test_orelse_short_circuit() {
    let output = run_vb_and_collect(r#"
Sub Main()
    Console.WriteLine(True OrElse True)
    Console.WriteLine(True OrElse False)
    Console.WriteLine(False OrElse True)
    Console.WriteLine(False OrElse False)
End Sub
"#);
    assert_eq!(output, vec!["True", "True", "True", "False"]);
}

#[test]
fn test_and_bitwise_on_integers() {
    let output = run_vb_and_collect(r#"
Sub Main()
    ' 12 = 1100, 10 = 1010, AND = 1000 = 8
    Console.WriteLine(12 And 10)
    ' 255 And 15 = 15
    Console.WriteLine(255 And 15)
End Sub
"#);
    assert_eq!(output, vec!["8", "15"]);
}

#[test]
fn test_or_bitwise_on_integers() {
    let output = run_vb_and_collect(r#"
Sub Main()
    ' 12 = 1100, 3 = 0011, OR = 1111 = 15
    Console.WriteLine(12 Or 3)
End Sub
"#);
    assert_eq!(output, vec!["15"]);
}

#[test]
fn test_xor_bitwise() {
    let output = run_vb_and_collect(r#"
Sub Main()
    ' 12 = 1100, 10 = 1010, XOR = 0110 = 6
    Console.WriteLine(12 Xor 10)
End Sub
"#);
    assert_eq!(output, vec!["6"]);
}

// ============================================================
// Bit Shifts
// ============================================================

#[test]
fn test_bit_shift_left() {
    let output = run_vb_and_collect(r#"
Sub Main()
    Console.WriteLine(1 << 0)
    Console.WriteLine(1 << 1)
    Console.WriteLine(1 << 4)
    Console.WriteLine(1 << 8)
End Sub
"#);
    assert_eq!(output, vec!["1", "2", "16", "256"]);
}

#[test]
fn test_bit_shift_right() {
    let output = run_vb_and_collect(r#"
Sub Main()
    Console.WriteLine(256 >> 1)
    Console.WriteLine(256 >> 4)
    Console.WriteLine(128 >> 3)
End Sub
"#);
    assert_eq!(output, vec!["128", "16", "16"]);
}

// ============================================================
// Comparison Operators
// ============================================================

#[test]
fn test_comparison_operators() {
    let output = run_vb_and_collect(r#"
Sub Main()
    Console.WriteLine(5 > 3)
    Console.WriteLine(3 > 5)
    Console.WriteLine(5 < 3)
    Console.WriteLine(3 < 5)
    Console.WriteLine(5 >= 5)
    Console.WriteLine(5 <= 5)
    Console.WriteLine(5 = 5)
    Console.WriteLine(5 <> 3)
    Console.WriteLine(5 <> 5)
End Sub
"#);
    assert_eq!(output, vec!["True", "False", "False", "True", "True", "True", "True", "True", "False"]);
}

// ============================================================
// String Concatenation
// ============================================================

#[test]
fn test_string_concatenation() {
    let output = run_vb_and_collect(r#"
Sub Main()
    Console.WriteLine("Hello" & " " & "World")
    Console.WriteLine("A" & "B" & "C" & "D")
    Dim x As Integer = 42
    Console.WriteLine("Value: " & x)
End Sub
"#);
    assert_eq!(output[0], "Hello World");
    assert_eq!(output[1], "ABCD");
    assert_eq!(output[2], "Value: 42");
}

// ============================================================
// Not Operator
// ============================================================

#[test]
fn test_not_operator() {
    let output = run_vb_and_collect(r#"
Sub Main()
    Console.WriteLine(Not True)
    Console.WriteLine(Not False)
End Sub
"#);
    assert_eq!(output, vec!["False", "True"]);
}
