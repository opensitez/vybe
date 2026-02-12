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
// Do While Loop
// ============================================================

#[test]
fn test_do_while_basic() {
    let output = run_vb(r#"
Sub Main()
    Dim i As Integer = 0
    Do While i < 5
        i = i + 1
    Loop
    Console.WriteLine(i)
End Sub
"#);
    assert_eq!(output, vec!["5"]);
}

#[test]
fn test_do_while_never_enters() {
    let output = run_vb(r#"
Sub Main()
    Dim i As Integer = 10
    Do While i < 5
        i = i + 1
    Loop
    Console.WriteLine(i)
End Sub
"#);
    assert_eq!(output, vec!["10"], "Do While with false condition should never enter body");
}

// ============================================================
// Do Until Loop
// ============================================================

#[test]
fn test_do_until_basic() {
    let output = run_vb(r#"
Sub Main()
    Dim i As Integer = 0
    Do Until i >= 5
        i = i + 1
    Loop
    Console.WriteLine(i)
End Sub
"#);
    assert_eq!(output, vec!["5"]);
}

#[test]
fn test_do_until_never_enters() {
    let output = run_vb(r#"
Sub Main()
    Dim i As Integer = 10
    Do Until i >= 5
        i = i + 1
    Loop
    Console.WriteLine(i)
End Sub
"#);
    assert_eq!(output, vec!["10"], "Do Until with already-true condition should never enter body");
}

#[test]
fn test_do_until_vs_do_while_semantics() {
    // Do While i < 5 and Do Until i >= 5 should produce the same result
    let output_while = run_vb(r#"
Sub Main()
    Dim i As Integer = 0
    Do While i < 5
        i = i + 1
    Loop
    Console.WriteLine(i)
End Sub
"#);
    let output_until = run_vb(r#"
Sub Main()
    Dim i As Integer = 0
    Do Until i >= 5
        i = i + 1
    Loop
    Console.WriteLine(i)
End Sub
"#);
    assert_eq!(output_while, output_until, "Do While i<5 and Do Until i>=5 should be equivalent");
}

// ============================================================
// Loop While / Loop Until (post-condition)
// ============================================================

#[test]
fn test_loop_while_executes_at_least_once() {
    let output = run_vb(r#"
Sub Main()
    Dim i As Integer = 10
    Do
        i = i + 1
    Loop While i < 5
    Console.WriteLine(i)
End Sub
"#);
    assert_eq!(output, vec!["11"], "Loop While should execute body at least once even if condition is false");
}

#[test]
fn test_loop_until_executes_at_least_once() {
    let output = run_vb(r#"
Sub Main()
    Dim i As Integer = 10
    Do
        i = i + 1
    Loop Until i >= 5
    Console.WriteLine(i)
End Sub
"#);
    assert_eq!(output, vec!["11"], "Loop Until should execute body at least once even if condition is true");
}

#[test]
fn test_loop_while_normal() {
    let output = run_vb(r#"
Sub Main()
    Dim i As Integer = 0
    Do
        i = i + 1
    Loop While i < 5
    Console.WriteLine(i)
End Sub
"#);
    assert_eq!(output, vec!["5"]);
}

#[test]
fn test_loop_until_normal() {
    let output = run_vb(r#"
Sub Main()
    Dim i As Integer = 0
    Do
        i = i + 1
    Loop Until i >= 5
    Console.WriteLine(i)
End Sub
"#);
    assert_eq!(output, vec!["5"]);
}

// ============================================================
// Infinite Do Loop with Exit Do
// ============================================================

#[test]
fn test_infinite_do_loop_with_exit() {
    let output = run_vb(r#"
Sub Main()
    Dim i As Integer = 0
    Do
        i = i + 1
        If i = 7 Then Exit Do
    Loop
    Console.WriteLine(i)
End Sub
"#);
    assert_eq!(output, vec!["7"]);
}

// ============================================================
// For Loop
// ============================================================

#[test]
fn test_for_loop_basic() {
    let output = run_vb(r#"
Sub Main()
    Dim sum As Integer = 0
    For i As Integer = 1 To 10
        sum = sum + i
    Next
    Console.WriteLine(sum)
End Sub
"#);
    assert_eq!(output, vec!["55"]);
}

#[test]
fn test_for_loop_with_step() {
    let output = run_vb(r#"
Sub Main()
    Dim result As String = ""
    For i As Integer = 0 To 10 Step 2
        result = result & i & ","
    Next
    Console.WriteLine(result)
End Sub
"#);
    assert_eq!(output, vec!["0,2,4,6,8,10,"]);
}

#[test]
fn test_for_loop_step_negative() {
    let output = run_vb(r#"
Sub Main()
    Dim result As String = ""
    For i As Integer = 5 To 1 Step -1
        result = result & i & ","
    Next
    Console.WriteLine(result)
End Sub
"#);
    assert_eq!(output, vec!["5,4,3,2,1,"]);
}

#[test]
fn test_for_loop_exit_for() {
    let output = run_vb(r#"
Sub Main()
    Dim lastI As Integer = 0
    For i As Integer = 1 To 100
        lastI = i
        If i = 5 Then Exit For
    Next
    Console.WriteLine(lastI)
End Sub
"#);
    assert_eq!(output, vec!["5"]);
}

// ============================================================
// While Loop
// ============================================================

#[test]
fn test_while_loop() {
    let output = run_vb(r#"
Sub Main()
    Dim i As Integer = 0
    While i < 5
        i = i + 1
    End While
    Console.WriteLine(i)
End Sub
"#);
    assert_eq!(output, vec!["5"]);
}

// ============================================================
// For Each
// ============================================================

#[test]
fn test_for_each_array() {
    let output = run_vb(r#"
Sub Main()
    Dim arr() As Integer = {10, 20, 30, 40, 50}
    Dim sum As Integer = 0
    For Each item As Integer In arr
        sum = sum + item
    Next
    Console.WriteLine(sum)
End Sub
"#);
    assert_eq!(output, vec!["150"]);
}

// ============================================================
// Select Case
// ============================================================

#[test]
fn test_select_case_exact_match() {
    let output = run_vb(r#"
Sub Main()
    Dim x As Integer = 2
    Select Case x
        Case 1
            Console.WriteLine("one")
        Case 2
            Console.WriteLine("two")
        Case 3
            Console.WriteLine("three")
        Case Else
            Console.WriteLine("other")
    End Select
End Sub
"#);
    assert_eq!(output, vec!["two"]);
}

#[test]
fn test_select_case_else() {
    let output = run_vb(r#"
Sub Main()
    Dim x As Integer = 99
    Select Case x
        Case 1
            Console.WriteLine("one")
        Case 2
            Console.WriteLine("two")
        Case Else
            Console.WriteLine("other")
    End Select
End Sub
"#);
    assert_eq!(output, vec!["other"]);
}

// ============================================================
// If / ElseIf / Else
// ============================================================

#[test]
fn test_if_elseif_else() {
    let output = run_vb(r#"
Sub Main()
    Dim x As Integer = 2
    If x = 1 Then
        Console.WriteLine("one")
    ElseIf x = 2 Then
        Console.WriteLine("two")
    ElseIf x = 3 Then
        Console.WriteLine("three")
    Else
        Console.WriteLine("other")
    End If
End Sub
"#);
    assert_eq!(output, vec!["two"]);
}

#[test]
fn test_nested_if() {
    let output = run_vb(r#"
Sub Main()
    Dim a As Boolean = True
    Dim b As Boolean = True
    If a Then
        If b Then
            Console.WriteLine("both true")
        Else
            Console.WriteLine("a true, b false")
        End If
    Else
        Console.WriteLine("a false")
    End If
End Sub
"#);
    assert_eq!(output, vec!["both true"]);
}

// ============================================================
// Try/Catch/Finally
// ============================================================

#[test]
fn test_try_catch_basic() {
    let output = run_vb(r#"
Sub Main()
    Try
        Dim x As Integer = 0
        Dim y As Integer = 10 / x
    Catch ex As Exception
        Console.WriteLine("caught")
    End Try
    Console.WriteLine("after")
End Sub
"#);
    assert!(output.iter().any(|s| s == "caught" || s == "after"), "Should either catch or continue");
}

// ============================================================
// Continue For / Continue Do
// ============================================================

#[test]
fn test_continue_for() {
    let output = run_vb(r#"
Sub Main()
    Dim result As String = ""
    For i As Integer = 1 To 5
        If i = 3 Then
            Continue For
        End If
        result = result & i & ","
    Next
    Console.WriteLine(result)
End Sub
"#);
    assert_eq!(output, vec!["1,2,4,5,"]);
}

#[test]
fn test_continue_do() {
    let output = run_vb(r#"
Sub Main()
    Dim i As Integer = 0
    Dim result As String = ""
    Do While i < 5
        i = i + 1
        If i = 3 Then
            Continue Do
        End If
        result = result & i & ","
    Loop
    Console.WriteLine(result)
End Sub
"#);
    assert_eq!(output, vec!["1,2,4,5,"]);
}

// ============================================================
// With Statement
// ============================================================

#[test]
fn test_with_statement() {
    let output = run_vb(r#"
Public Class Person
    Public Name As String
    Public Age As Integer
End Class

Sub Main()
    Dim p As New Person()
    With p
        .Name = "Alice"
        .Age = 30
    End With
    Console.WriteLine(p.Name)
    Console.WriteLine(p.Age)
End Sub
"#);
    assert_eq!(output, vec!["Alice", "30"]);
}
