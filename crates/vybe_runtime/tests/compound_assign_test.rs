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

// ============================================================
// += (AddAssign)
// ============================================================

#[test]
fn test_add_assign_integer() {
    let output = run_vb(r#"
Sub Main()
    Dim x As Integer = 5
    x += 3
    Console.WriteLine(x)
End Sub
"#);
    assert_eq!(output, vec!["8"]);
}

#[test]
fn test_add_assign_double() {
    let output = run_vb(r#"
Sub Main()
    Dim x As Double = 1.5
    x += 2.5
    Console.WriteLine(x)
End Sub
"#);
    assert_eq!(output, vec!["4"]);
}

// ============================================================
// -= (SubtractAssign)
// ============================================================

#[test]
fn test_subtract_assign() {
    let output = run_vb(r#"
Sub Main()
    Dim x As Integer = 10
    x -= 3
    Console.WriteLine(x)
End Sub
"#);
    assert_eq!(output, vec!["7"]);
}

// ============================================================
// *= (MultiplyAssign)
// ============================================================

#[test]
fn test_multiply_assign() {
    let output = run_vb(r#"
Sub Main()
    Dim x As Integer = 5
    x *= 4
    Console.WriteLine(x)
End Sub
"#);
    assert_eq!(output, vec!["20"]);
}

// ============================================================
// /= (DivideAssign)
// ============================================================

#[test]
fn test_divide_assign() {
    let output = run_vb(r#"
Sub Main()
    Dim x As Double = 20.0
    x /= 4.0
    Console.WriteLine(x)
End Sub
"#);
    assert_eq!(output, vec!["5"]);
}

// ============================================================
// \= (IntDivideAssign)
// ============================================================

#[test]
fn test_int_divide_assign() {
    let output = run_vb(r#"
Sub Main()
    Dim x As Integer = 17
    x \= 5
    Console.WriteLine(x)
End Sub
"#);
    assert_eq!(output, vec!["3"]);
}

// ============================================================
// &= (ConcatAssign)
// ============================================================

#[test]
fn test_concat_assign() {
    let output = run_vb(r#"
Sub Main()
    Dim s As String = "Hello"
    s &= " World"
    Console.WriteLine(s)
End Sub
"#);
    assert_eq!(output, vec!["Hello World"]);
}

#[test]
fn test_concat_assign_multiple() {
    let output = run_vb(r#"
Sub Main()
    Dim s As String = ""
    s &= "A"
    s &= "B"
    s &= "C"
    Console.WriteLine(s)
End Sub
"#);
    assert_eq!(output, vec!["ABC"]);
}

// ============================================================
// ^= (ExponentAssign)
// ============================================================

#[test]
fn test_exponent_assign() {
    let output = run_vb(r#"
Sub Main()
    Dim x As Double = 2.0
    x ^= 10
    Console.WriteLine(x)
End Sub
"#);
    assert_eq!(output, vec!["1024"]);
}

// ============================================================
// <<= (ShiftLeftAssign)
// ============================================================

#[test]
fn test_shift_left_assign() {
    let output = run_vb(r#"
Sub Main()
    Dim x As Integer = 1
    x <<= 4
    Console.WriteLine(x)
End Sub
"#);
    assert_eq!(output, vec!["16"]);
}

// ============================================================
// >>= (ShiftRightAssign)
// ============================================================

#[test]
fn test_shift_right_assign() {
    let output = run_vb(r#"
Sub Main()
    Dim x As Integer = 128
    x >>= 3
    Console.WriteLine(x)
End Sub
"#);
    assert_eq!(output, vec!["16"]);
}

// ============================================================
// Compound assignment in loops
// ============================================================

#[test]
fn test_compound_assign_in_loop() {
    let output = run_vb(r#"
Sub Main()
    Dim sum As Integer = 0
    For i As Integer = 1 To 10
        sum += i
    Next
    Console.WriteLine(sum)
End Sub
"#);
    assert_eq!(output, vec!["55"]);
}

#[test]
fn test_compound_string_concat_in_loop() {
    let output = run_vb(r#"
Sub Main()
    Dim result As String = ""
    For i As Integer = 1 To 5
        result &= CStr(i)
    Next
    Console.WriteLine(result)
End Sub
"#);
    assert_eq!(output, vec!["12345"]);
}

// ============================================================
// Mixed compound assignments
// ============================================================

#[test]
fn test_all_compound_ops_sequence() {
    let output = run_vb(r#"
Sub Main()
    Dim x As Double = 10.0
    x += 5      ' 15
    Console.WriteLine(x)
    x -= 3      ' 12
    Console.WriteLine(x)
    x *= 2      ' 24
    Console.WriteLine(x)
    x /= 4      ' 6
    Console.WriteLine(x)
    x ^= 2      ' 36
    Console.WriteLine(x)
End Sub
"#);
    assert_eq!(output, vec!["15", "12", "24", "6", "36"]);
}
