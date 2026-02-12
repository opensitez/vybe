///! Comprehensive tests for array operations including ReDim, ReDim Preserve,
///! array initialization, access, modification, and edge cases.

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
// Basic Array Init and Access
// ============================================================

#[test]
fn test_array_literal_init() {
    let output = run_vb(r#"
Sub Main()
    Dim arr() As Integer = {10, 20, 30, 40, 50}
    Console.WriteLine(arr(0))
    Console.WriteLine(arr(2))
    Console.WriteLine(arr(4))
End Sub
"#);
    assert_eq!(output, vec!["10", "30", "50"]);
}

#[test]
fn test_array_modification() {
    let output = run_vb(r#"
Sub Main()
    Dim arr() As Integer = {1, 2, 3}
    arr(1) = 99
    Console.WriteLine(arr(0) & "," & arr(1) & "," & arr(2))
End Sub
"#);
    assert_eq!(output, vec!["1,99,3"]);
}

// ============================================================
// ReDim Preserve — Grow
// ============================================================

#[test]
fn test_redim_preserve_grow() {
    // Grow array from 3 to 6, verify old data preserved
    let output = run_vb(r#"
Sub Main()
    Dim arr() As Integer = {1, 2, 3}
    ReDim Preserve arr(5)
    arr(3) = 4
    arr(4) = 5
    arr(5) = 6
    Console.WriteLine(arr(0) & "," & arr(1) & "," & arr(2))
    Console.WriteLine(arr(3) & "," & arr(4) & "," & arr(5))
End Sub
"#);
    assert_eq!(output[0], "1,2,3");
    assert_eq!(output[1], "4,5,6");
}

#[test]
fn test_redim_preserve_grow_from_empty() {
    // Start with empty array, grow with Preserve
    let output = run_vb(r#"
Sub Main()
    Dim arr() As Integer = {}
    ReDim Preserve arr(2)
    arr(0) = 10
    arr(1) = 20
    arr(2) = 30
    Console.WriteLine(arr(0) & "," & arr(1) & "," & arr(2))
End Sub
"#);
    assert_eq!(output, vec!["10,20,30"]);
}

#[test]
fn test_redim_preserve_grow_double() {
    // Grow twice with Preserve
    let output = run_vb(r#"
Sub Main()
    Dim arr() As Integer = {1, 2}
    ReDim Preserve arr(3)
    arr(2) = 3
    arr(3) = 4
    ReDim Preserve arr(5)
    arr(4) = 5
    arr(5) = 6
    Console.WriteLine(arr(0) & "," & arr(1) & "," & arr(2) & "," & arr(3) & "," & arr(4) & "," & arr(5))
End Sub
"#);
    assert_eq!(output, vec!["1,2,3,4,5,6"]);
}

// ============================================================
// ReDim Preserve — Shrink
// ============================================================

#[test]
fn test_redim_preserve_shrink() {
    // Shrink array from 5 to 3, verify remaining data
    let output = run_vb(r#"
Sub Main()
    Dim arr() As Integer = {10, 20, 30, 40, 50}
    ReDim Preserve arr(2)
    Console.WriteLine(arr(0) & "," & arr(1) & "," & arr(2))
End Sub
"#);
    assert_eq!(output, vec!["10,20,30"]);
}

// ============================================================
// ReDim Without Preserve — Resets Array
// ============================================================

#[test]
fn test_redim_without_preserve() {
    // ReDim without Preserve resets the array to default values
    let output = run_vb(r#"
Sub Main()
    Dim arr() As Integer = {1, 2, 3}
    ReDim arr(4)
    ' New elements should be 0 (default for Integer)
    Console.WriteLine(arr(0))
    Console.WriteLine(arr(4))
    arr(0) = 100
    arr(4) = 500
    Console.WriteLine(arr(0) & "," & arr(4))
End Sub
"#);
    assert_eq!(output[0], "0");
    assert_eq!(output[1], "0");
    assert_eq!(output[2], "100,500");
}

// ============================================================
// ReDim Preserve with String Arrays
// ============================================================

#[test]
fn test_redim_preserve_string_array() {
    let output = run_vb(r#"
Sub Main()
    Dim names() As String = {"Alice", "Bob"}
    ReDim Preserve names(3)
    names(2) = "Charlie"
    names(3) = "Diana"
    Console.WriteLine(names(0) & "," & names(1))
    Console.WriteLine(names(2) & "," & names(3))
End Sub
"#);
    assert_eq!(output[0], "Alice,Bob");
    assert_eq!(output[1], "Charlie,Diana");
}

// ============================================================
// ReDim Preserve in Loop (dynamic growth)
// ============================================================

#[test]
fn test_redim_preserve_in_loop() {
    // Build array dynamically in a loop
    let output = run_vb(r#"
Sub Main()
    Dim arr() As Integer = {}
    For i As Integer = 0 To 4
        ReDim Preserve arr(i)
        arr(i) = (i + 1) * 10
    Next
    Console.WriteLine(arr(0) & "," & arr(1) & "," & arr(2) & "," & arr(3) & "," & arr(4))
End Sub
"#);
    assert_eq!(output, vec!["10,20,30,40,50"]);
}

// ============================================================
// ReDim Preserve — Verify Size (UBound equivalent)
// ============================================================

#[test]
fn test_redim_preserve_size_after_grow() {
    // After ReDim Preserve arr(9), there should be 10 elements (0..9)
    let output = run_vb(r#"
Sub Main()
    Dim arr() As Integer = {1, 2, 3}
    ReDim Preserve arr(9)
    arr(9) = 999
    Console.WriteLine(arr(0))
    Console.WriteLine(arr(2))
    Console.WriteLine(arr(9))
End Sub
"#);
    assert_eq!(output, vec!["1", "3", "999"]);
}

// ============================================================
// ReDim Preserve — Preserves Type Coercion
// ============================================================

#[test]
fn test_redim_preserve_double_array() {
    let output = run_vb(r#"
Sub Main()
    Dim vals() As Double = {1.1, 2.2, 3.3}
    ReDim Preserve vals(4)
    vals(3) = 4.4
    vals(4) = 5.5
    Console.WriteLine(vals(0))
    Console.WriteLine(vals(3))
    Console.WriteLine(vals(4))
End Sub
"#);
    assert!(output[0].starts_with("1.1"));
    assert!(output[1].starts_with("4.4"));
    assert!(output[2].starts_with("5.5"));
}

// ============================================================
// ReDim to size 0 (single element)
// ============================================================

#[test]
fn test_redim_to_zero() {
    // ReDim arr(0) means 1 element at index 0
    let output = run_vb(r#"
Sub Main()
    Dim arr() As Integer = {1, 2, 3, 4, 5}
    ReDim arr(0)
    arr(0) = 42
    Console.WriteLine(arr(0))
End Sub
"#);
    assert_eq!(output, vec!["42"]);
}

// ============================================================
// ReDim Preserve with computed bound
// ============================================================

#[test]
fn test_redim_preserve_computed_bound() {
    let output = run_vb(r#"
Sub Main()
    Dim arr() As Integer = {1, 2, 3}
    Dim newSize As Integer = 3 + 3
    ReDim Preserve arr(newSize)
    arr(newSize) = 77
    Console.WriteLine(arr(0) & "," & arr(1) & "," & arr(2))
    Console.WriteLine(arr(newSize))
End Sub
"#);
    assert_eq!(output[0], "1,2,3");
    assert_eq!(output[1], "77");
}

// ============================================================
// Parser: ReDim and ReDim Preserve produce correct AST
// ============================================================

#[test]
fn test_redim_parses() {
    use irys_parser::parse_program;
    use irys_parser::ast::stmt::Statement;
    use irys_parser::ast::decl::Declaration;

    let code = r#"
Sub Main()
    Dim arr() As Integer = {1, 2, 3}
    ReDim Preserve arr(10)
    ReDim arr(5)
End Sub
"#;
    let prog = parse_program(code).expect("Parse error");
    let stmts: Vec<_> = prog.declarations.iter().filter_map(|d| {
        if let Declaration::Sub(s) = d { Some(s.body.clone()) } else { None }
    }).flatten().collect();

    // Find ReDim statements
    let redims: Vec<_> = stmts.iter().filter(|s| matches!(s, Statement::ReDim { .. })).collect();
    assert_eq!(redims.len(), 2, "Should have 2 ReDim statements");

    if let Statement::ReDim { preserve, array, bounds } = &redims[0] {
        assert!(*preserve, "First ReDim should have Preserve");
        assert_eq!(array.as_str(), "arr");
        assert_eq!(bounds.len(), 1);
    }

    if let Statement::ReDim { preserve, array, bounds } = &redims[1] {
        assert!(!preserve, "Second ReDim should NOT have Preserve");
        assert_eq!(array.as_str(), "arr");
        assert_eq!(bounds.len(), 1);
    }
}
