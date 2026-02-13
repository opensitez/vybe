
use vybe_parser::parser::parse_program;
// use vybe_parser::ast::*;
use vybe_parser::ast::Declaration;

// Test for single argument implicit call
#[test]
fn test_implicit_call_one_arg() {
    let code = r#"
        Sub Test()
            MsgBox "Hello"
        End Sub
    "#;
    let result = parse_program(code);
    assert!(result.is_ok(), "Failed to parse: {:?}", result.err());
}

// Test for multiple arguments implicit call
#[test]
fn test_implicit_call_multi_args() {
    let code = r#"
        Sub Test()
            Foo Arg1, 42
        End Sub
    "#;
    let result = parse_program(code);
    assert!(result.is_ok(), "Failed to parse: {:?}", result.err());
}

// Test for call with parens (explicit call)
#[test]
fn test_explicit_call_parens() {
    let code = r#"
        Sub Test()
            Call MsgBox("Hello")
        End Sub
    "#;
    let result = parse_program(code);
    assert!(result.is_ok(), "Failed to parse: {:?}", result.err());
}

// Test call without args
#[test]
fn test_call_no_args() {
    let code = r#"
        Sub Test()
            DoSomething
        End Sub
    "#;
    let result = parse_program(code);
    assert!(result.is_ok(), "Failed to parse: {:?}", result.err());
}

#[test]
fn test_assignment() {
    let code = r#"
        Sub Test()
            x = 1
        End Sub
    "#;
    let result = parse_program(code);
    assert!(result.is_ok(), "Failed to parse assignment: {:?}", result.err());
}

#[test]
fn test_reproduction_empty_lines() {
    let input = r#"
Sub Test()
    MsgBox "Hello"
End Sub

"#; 
    // The input above has a trailing newline which might be line 5 if counted.
    // Line 1: empty (if starts with newline) or Sub Test
    // Let's try to match the user's "5:1" error.
    // 1: Sub Test()
    // 2:     MsgBox "Hello"
    // 3: End Sub
    // 4: 
    // 5: <EOF>
    
    let result = vybe_parser::parse_program(input);
    assert!(result.is_ok(), "Failed to parse program with trailing newlines: {:?}", result.err());
}

#[test]
fn test_trailing_newlines() {
    let code = "Private Sub btn_Click()\n    MsgBox \"Hi\"\nEnd Sub\n\n\n";
    let result = parse_program(code);
    assert!(result.is_ok(), "Failed to parse with trailing newlines: {:?}", result.err());
}
#[test]
fn test_repro_user_errors() {
    // Case 1: Empty line between sub header and body
    let code1 = "Sub btn1_Click()\n\n    MsgBox \"Hi\"\nEnd Sub";
    assert!(parse_program(code1).is_ok(), "Failed code1");

    // Case 2: Empty line before End Sub
    let code2 = "Sub btn1_Click()\n    MsgBox \"Hi\"\n\nEnd Sub";
    assert!(parse_program(code2).is_ok(), "Failed code2");

    // Case 3: Spaces and newlines everywhere
    let code3 = " \n\nSub Test()  \n  \n  x = 1 \n \nEnd Sub \n \n ";
    assert!(parse_program(code3).is_ok(), "Failed code3");
}

#[test]
fn test_fifty_empty_lines() {
    let mut code = String::from("Sub Test()\n");
    for _ in 0..50 {
        code.push_str("   \n\t\n\n");
    }
    code.push_str("    MsgBox \"Done\"\n");
    for _ in 0..50 {
        code.push_str("   \n\t\n\n");
    }
    code.push_str("End Sub");
    
    let result = parse_program(&code);
    assert!(result.is_ok(), "Failed to parse with 50+ empty lines: {:?}", result.err());
}

#[test]
fn test_interpolated_string_basic() {
    let code = r#"
        Sub Test()
            Dim name As String = "World"
            Dim result As String = $"Hello {name}!"
        End Sub
    "#;
    let result = parse_program(code);
    assert!(result.is_ok(), "Failed to parse basic interpolated string: {:?}", result.err());
}

#[test]
fn test_interpolated_string_expression() {
    let code = r#"
        Sub Test()
            Dim x As Integer = 5
            Dim y As Integer = 3
            Dim result As String = $"Sum is {x + y}"
        End Sub
    "#;
    let result = parse_program(code);
    assert!(result.is_ok(), "Failed to parse interpolated string with expression: {:?}", result.err());
}

#[test]
fn test_interpolated_string_multiple() {
    let code = r#"
        Sub Test()
            Dim a As String = "A"
            Dim b As String = "B"
            Dim result As String = $"{a} and {b}"
        End Sub
    "#;
    let result = parse_program(code);
    assert!(result.is_ok(), "Failed to parse interpolated string with multiple exprs: {:?}", result.err());
}

#[test]
fn test_interpolated_string_method_call() {
    let code = r#"
        Sub Test()
            Dim s As String = "hello"
            Dim result As String = $"Upper: {s.ToUpper()}"
        End Sub
    "#;
    let result = parse_program(code);
    assert!(result.is_ok(), "Failed to parse interpolated string with method call: {:?}", result.err());
}

#[test]
fn test_if_expression_ternary() {
    let code = r#"
        Sub Test()
            Dim x As Integer = If(True, 1, 0)
        End Sub
    "#;
    let result = parse_program(code);
    assert!(result.is_ok(), "Failed to parse ternary If expression: {:?}", result.err());
}

#[test]
fn test_if_expression_coalesce() {
    let code = r#"
        Sub Test()
            Dim x As String = If(Nothing, "default")
        End Sub
    "#;
    let result = parse_program(code);
    assert!(result.is_ok(), "Failed to parse coalesce If expression: {:?}", result.err());
}

#[test]
fn test_nullable_type_param() {
    let code = r#"
        Function Test(Optional timeout? As Integer = Nothing) As String
            Return "ok"
        End Function
    "#;
    let result = parse_program(code);
    assert!(result.is_ok(), "Failed to parse nullable type param: {:?}", result.err());
}

// ===== Extension method tests =====

#[test]
fn test_extension_function_attribute() {
    let code = r#"
        <Extension()>
        Function Reverse(s As String) As String
            Return StrReverse(s)
        End Function
    "#;
    let prog = parse_program(code).expect("Failed to parse extension function");
    let func = prog.declarations.iter().find_map(|d| {
        if let Declaration::Function(f) = d { Some(f) } else { None }
    }).expect("No function declaration found");
    assert!(func.is_extension, "Function should be marked as extension");
    assert_eq!(func.name.as_str(), "Reverse");
}

#[test]
fn test_extension_sub_attribute() {
    let code = r#"
        <Extension()>
        Sub PrintUpper(s As String)
            Console.WriteLine(s.ToUpper())
        End Sub
    "#;
    let prog = parse_program(code).expect("Failed to parse extension sub");
    let sub_decl = prog.declarations.iter().find_map(|d| {
        if let Declaration::Sub(s) = d { Some(s) } else { None }
    }).expect("No sub declaration found");
    assert!(sub_decl.is_extension, "Sub should be marked as extension");
    assert_eq!(sub_decl.name.as_str(), "PrintUpper");
}

#[test]
fn test_extension_qualified_attribute() {
    let code = r#"
        <Runtime.CompilerServices.Extension()>
        Function IsNullOrEmpty(s As String) As Boolean
            Return s Is Nothing OrElse s = ""
        End Function
    "#;
    let prog = parse_program(code).expect("Failed to parse qualified extension attribute");
    let func = prog.declarations.iter().find_map(|d| {
        if let Declaration::Function(f) = d { Some(f) } else { None }
    }).expect("No function declaration found");
    assert!(func.is_extension, "Function with qualified attribute should be marked as extension");
}

#[test]
fn test_non_extension_function() {
    let code = r#"
        Function NormalFunc(x As Integer) As Integer
            Return x + 1
        End Function
    "#;
    let prog = parse_program(code).expect("Failed to parse normal function");
    let func = prog.declarations.iter().find_map(|d| {
        if let Declaration::Function(f) = d { Some(f) } else { None }
    }).expect("No function declaration found");
    assert!(!func.is_extension, "Normal function should NOT be marked as extension");
}
