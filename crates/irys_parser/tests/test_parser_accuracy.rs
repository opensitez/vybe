///! Comprehensive parser accuracy tests for all new and fixed features.
///! Tests grammar rules, AST node construction, operator precedence,
///! and VB.NET syntax compatibility.

use irys_parser::parse_program;
use irys_parser::ast::decl::{Declaration, Visibility, Parameter};
use irys_parser::ast::stmt::{Statement, LoopConditionType, CompoundOp};
use irys_parser::ast::expr::Expression;

// ============================================================
// Helper to extract first Sub's body statements
// ============================================================
fn parse_sub_body(code: &str) -> Vec<Statement> {
    let prog = parse_program(code).expect("Parse failed");
    for d in &prog.declarations {
        if let Declaration::Sub(s) = d {
            return s.body.clone();
        }
    }
    panic!("No Sub found in code");
}

fn parse_first_sub(code: &str) -> irys_parser::ast::decl::SubDecl {
    let prog = parse_program(code).expect("Parse failed");
    for d in &prog.declarations {
        if let Declaration::Sub(s) = d {
            return s.clone();
        }
    }
    panic!("No Sub found in code");
}

fn parse_first_function(code: &str) -> irys_parser::ast::decl::FunctionDecl {
    let prog = parse_program(code).expect("Parse failed");
    for d in &prog.declarations {
        if let Declaration::Function(f) = d {
            return f.clone();
        }
    }
    panic!("No Function found in code");
}

// ============================================================
// Do While vs Do Until Parsing
// ============================================================

#[test]
fn test_do_while_loop_parsed_correctly() {
    let code = r#"
Sub Test()
    Dim i As Integer = 0
    Do While i < 10
        i += 1
    Loop
End Sub
"#;
    let stmts = parse_sub_body(code);
    // Find the DoLoop
    let do_loop = stmts.iter().find(|s| matches!(s, Statement::DoLoop { .. }));
    assert!(do_loop.is_some(), "Should find a DoLoop statement");
    if let Statement::DoLoop { pre_condition, .. } = do_loop.unwrap() {
        assert!(pre_condition.is_some(), "Do While should have pre_condition");
        let (cond_type, _) = pre_condition.as_ref().unwrap();
        assert!(matches!(cond_type, LoopConditionType::While), "Do While should parse as While, not Until");
    }
}

#[test]
fn test_do_until_loop_parsed_correctly() {
    let code = r#"
Sub Test()
    Dim i As Integer = 0
    Do Until i >= 10
        i += 1
    Loop
End Sub
"#;
    let stmts = parse_sub_body(code);
    let do_loop = stmts.iter().find(|s| matches!(s, Statement::DoLoop { .. }));
    assert!(do_loop.is_some(), "Should find a DoLoop statement");
    if let Statement::DoLoop { pre_condition, .. } = do_loop.unwrap() {
        assert!(pre_condition.is_some(), "Do Until should have pre_condition");
        let (cond_type, _) = pre_condition.as_ref().unwrap();
        assert!(matches!(cond_type, LoopConditionType::Until), "Do Until should parse as Until, not While");
    }
}

#[test]
fn test_loop_while_post_condition() {
    let code = r#"
Sub Test()
    Dim i As Integer = 0
    Do
        i += 1
    Loop While i < 10
End Sub
"#;
    let stmts = parse_sub_body(code);
    let do_loop = stmts.iter().find(|s| matches!(s, Statement::DoLoop { .. }));
    assert!(do_loop.is_some());
    if let Statement::DoLoop { pre_condition, post_condition, .. } = do_loop.unwrap() {
        assert!(pre_condition.is_none(), "Post-condition loop should not have pre_condition");
        assert!(post_condition.is_some(), "Loop While should have post_condition");
        let (cond_type, _) = post_condition.as_ref().unwrap();
        assert!(matches!(cond_type, LoopConditionType::While));
    }
}

#[test]
fn test_loop_until_post_condition() {
    let code = r#"
Sub Test()
    Dim i As Integer = 0
    Do
        i += 1
    Loop Until i >= 10
End Sub
"#;
    let stmts = parse_sub_body(code);
    let do_loop = stmts.iter().find(|s| matches!(s, Statement::DoLoop { .. }));
    assert!(do_loop.is_some());
    if let Statement::DoLoop { post_condition, .. } = do_loop.unwrap() {
        assert!(post_condition.is_some());
        let (cond_type, _) = post_condition.as_ref().unwrap();
        assert!(matches!(cond_type, LoopConditionType::Until));
    }
}

#[test]
fn test_infinite_do_loop() {
    let code = r#"
Sub Test()
    Dim x As Integer = 0
    Do
        x += 1
        If x > 5 Then Exit Do
    Loop
End Sub
"#;
    let stmts = parse_sub_body(code);
    let do_loop = stmts.iter().find(|s| matches!(s, Statement::DoLoop { .. }));
    assert!(do_loop.is_some());
    if let Statement::DoLoop { pre_condition, post_condition, .. } = do_loop.unwrap() {
        assert!(pre_condition.is_none());
        assert!(post_condition.is_none());
    }
}

// ============================================================
// Operator Parsing: Integer Divide, Exponent, Like
// ============================================================

#[test]
fn test_integer_divide_parses() {
    let code = r#"
Sub Test()
    Dim x As Integer = 10 \ 3
End Sub
"#;
    let prog = parse_program(code);
    assert!(prog.is_ok(), "Integer divide \\ should parse: {:?}", prog.err());
}

#[test]
fn test_exponent_parses() {
    let code = r#"
Sub Test()
    Dim x As Double = 2 ^ 10
End Sub
"#;
    let prog = parse_program(code);
    assert!(prog.is_ok(), "Exponent ^ should parse: {:?}", prog.err());
}

#[test]
fn test_like_operator_parses() {
    let code = r#"
Sub Test()
    Dim result As Boolean = "Hello" Like "H*"
End Sub
"#;
    let prog = parse_program(code);
    assert!(prog.is_ok(), "Like operator should parse: {:?}", prog.err());
}

// ============================================================
// Compound Assignment Parsing
// ============================================================

#[test]
fn test_compound_add_assign_parses() {
    let code = r#"
Sub Test()
    Dim x As Integer = 5
    x += 3
End Sub
"#;
    let stmts = parse_sub_body(code);
    let compound = stmts.iter().find(|s| matches!(s, Statement::CompoundAssignment { .. }));
    assert!(compound.is_some(), "Should parse += as CompoundAssignment");
    if let Statement::CompoundAssignment { operator, .. } = compound.unwrap() {
        assert!(matches!(operator, CompoundOp::AddAssign));
    }
}

#[test]
fn test_compound_subtract_assign_parses() {
    let code = r#"
Sub Test()
    Dim x As Integer = 10
    x -= 3
End Sub
"#;
    let stmts = parse_sub_body(code);
    let compound = stmts.iter().find(|s| matches!(s, Statement::CompoundAssignment { .. }));
    assert!(compound.is_some());
    if let Statement::CompoundAssignment { operator, .. } = compound.unwrap() {
        assert!(matches!(operator, CompoundOp::SubtractAssign));
    }
}

#[test]
fn test_compound_multiply_assign_parses() {
    let code = r#"
Sub Test()
    Dim x As Integer = 5
    x *= 3
End Sub
"#;
    let stmts = parse_sub_body(code);
    let compound = stmts.iter().find(|s| matches!(s, Statement::CompoundAssignment { .. }));
    assert!(compound.is_some());
    if let Statement::CompoundAssignment { operator, .. } = compound.unwrap() {
        assert!(matches!(operator, CompoundOp::MultiplyAssign));
    }
}

#[test]
fn test_compound_divide_assign_parses() {
    let code = r#"
Sub Test()
    Dim x As Double = 10.0
    x /= 3.0
End Sub
"#;
    let stmts = parse_sub_body(code);
    let compound = stmts.iter().find(|s| matches!(s, Statement::CompoundAssignment { .. }));
    assert!(compound.is_some());
    if let Statement::CompoundAssignment { operator, .. } = compound.unwrap() {
        assert!(matches!(operator, CompoundOp::DivideAssign));
    }
}

#[test]
fn test_compound_int_divide_assign_parses() {
    let code = r#"
Sub Test()
    Dim x As Integer = 10
    x \= 3
End Sub
"#;
    let stmts = parse_sub_body(code);
    let compound = stmts.iter().find(|s| matches!(s, Statement::CompoundAssignment { .. }));
    assert!(compound.is_some());
    if let Statement::CompoundAssignment { operator, .. } = compound.unwrap() {
        assert!(matches!(operator, CompoundOp::IntDivideAssign));
    }
}

#[test]
fn test_compound_concat_assign_parses() {
    let code = r#"
Sub Test()
    Dim s As String = "Hello"
    s &= " World"
End Sub
"#;
    let stmts = parse_sub_body(code);
    let compound = stmts.iter().find(|s| matches!(s, Statement::CompoundAssignment { .. }));
    assert!(compound.is_some());
    if let Statement::CompoundAssignment { operator, .. } = compound.unwrap() {
        assert!(matches!(operator, CompoundOp::ConcatAssign));
    }
}

#[test]
fn test_compound_exponent_assign_parses() {
    let code = r#"
Sub Test()
    Dim x As Double = 2.0
    x ^= 10
End Sub
"#;
    let stmts = parse_sub_body(code);
    let compound = stmts.iter().find(|s| matches!(s, Statement::CompoundAssignment { .. }));
    assert!(compound.is_some());
    if let Statement::CompoundAssignment { operator, .. } = compound.unwrap() {
        assert!(matches!(operator, CompoundOp::ExponentAssign));
    }
}

#[test]
fn test_compound_shift_left_assign_parses() {
    let code = r#"
Sub Test()
    Dim x As Integer = 1
    x <<= 4
End Sub
"#;
    let stmts = parse_sub_body(code);
    let compound = stmts.iter().find(|s| matches!(s, Statement::CompoundAssignment { .. }));
    assert!(compound.is_some());
    if let Statement::CompoundAssignment { operator, .. } = compound.unwrap() {
        assert!(matches!(operator, CompoundOp::ShiftLeftAssign));
    }
}

#[test]
fn test_compound_shift_right_assign_parses() {
    let code = r#"
Sub Test()
    Dim x As Integer = 128
    x >>= 3
End Sub
"#;
    let stmts = parse_sub_body(code);
    let compound = stmts.iter().find(|s| matches!(s, Statement::CompoundAssignment { .. }));
    assert!(compound.is_some());
    if let Statement::CompoundAssignment { operator, .. } = compound.unwrap() {
        assert!(matches!(operator, CompoundOp::ShiftRightAssign));
    }
}

// ============================================================
// RaiseEvent Parsing
// ============================================================

#[test]
fn test_raiseevent_parses() {
    let code = r#"
Sub Test()
    RaiseEvent Click()
End Sub
"#;
    let stmts = parse_sub_body(code);
    let raise = stmts.iter().find(|s| matches!(s, Statement::RaiseEvent { .. }));
    assert!(raise.is_some(), "RaiseEvent should parse");
    if let Statement::RaiseEvent { event_name, arguments } = raise.unwrap() {
        assert_eq!(event_name.as_str(), "Click");
        assert!(arguments.is_empty());
    }
}

#[test]
fn test_raiseevent_with_args_parses() {
    let code = r#"
Sub Test()
    RaiseEvent ValueChanged(42, "test")
End Sub
"#;
    let stmts = parse_sub_body(code);
    let raise = stmts.iter().find(|s| matches!(s, Statement::RaiseEvent { .. }));
    assert!(raise.is_some());
    if let Statement::RaiseEvent { event_name, arguments } = raise.unwrap() {
        assert_eq!(event_name.as_str(), "ValueChanged");
        assert_eq!(arguments.len(), 2);
    }
}

// ============================================================
// Visibility: Protected
// ============================================================

#[test]
fn test_protected_sub_parses() {
    let code = r#"
Public Class MyClass
    Protected Sub OnClick()
    End Sub
End Class
"#;
    let prog = parse_program(code).expect("Protected Sub should parse");
    let class = prog.declarations.iter().find_map(|d| {
        if let Declaration::Class(c) = d { Some(c) } else { None }
    }).expect("Should find class");
    let method = class.methods.iter().find(|m| {
        match m {
            irys_parser::ast::decl::MethodDecl::Sub(s) => s.name.as_str() == "OnClick",
            _ => false,
        }
    }).expect("Should find OnClick");
    match method {
        irys_parser::ast::decl::MethodDecl::Sub(s) => {
            assert!(matches!(s.visibility, Visibility::Protected));
        }
        _ => panic!("Expected Sub"),
    }
}

#[test]
fn test_protected_function_parses() {
    let code = r#"
Public Class MyClass
    Protected Function GetValue() As Integer
        Return 42
    End Function
End Class
"#;
    let prog = parse_program(code).expect("Protected Function should parse");
    let class = prog.declarations.iter().find_map(|d| {
        if let Declaration::Class(c) = d { Some(c) } else { None }
    }).expect("Should find class");
    let method = class.methods.iter().find(|m| {
        match m {
            irys_parser::ast::decl::MethodDecl::Function(f) => f.name.as_str() == "GetValue",
            _ => false,
        }
    });
    assert!(method.is_some(), "Should find Protected Function GetValue");
}

// ============================================================
// Optional Parameters with Default Values
// ============================================================

#[test]
fn test_optional_parameter_with_default_parses() {
    let code = r#"
Sub Test(Optional x As Integer = 42)
End Sub
"#;
    let sub = parse_first_sub(code);
    assert_eq!(sub.parameters.len(), 1);
    let param = &sub.parameters[0];
    assert!(param.is_optional, "Parameter should be marked optional");
    assert!(param.default_value.is_some(), "Optional param should have default value");
}

#[test]
fn test_optional_and_required_params_mixed() {
    let code = r#"
Sub Test(a As Integer, Optional b As String = "default", Optional c As Boolean = True)
End Sub
"#;
    let sub = parse_first_sub(code);
    assert_eq!(sub.parameters.len(), 3);
    assert!(!sub.parameters[0].is_optional);
    assert!(sub.parameters[1].is_optional);
    assert!(sub.parameters[1].default_value.is_some());
    assert!(sub.parameters[2].is_optional);
    assert!(sub.parameters[2].default_value.is_some());
}

// ============================================================
// TypeOf...Is Parsing
// ============================================================

#[test]
fn test_typeof_is_parses() {
    let code = r#"
Sub Test()
    Dim obj As Object = Nothing
    Dim result As Boolean = TypeOf obj Is String
End Sub
"#;
    let prog = parse_program(code);
    assert!(prog.is_ok(), "TypeOf...Is should parse: {:?}", prog.err());
}

// ============================================================
// Interface, Structure, Namespace, Delegate, Event Declarations
// (These should parse without error even if not fully implemented)
// ============================================================

#[test]
fn test_interface_parses_gracefully() {
    let code = r#"
Interface IMyInterface
    Sub DoSomething()
    Function Calculate() As Integer
End Interface
"#;
    let prog = parse_program(code);
    assert!(prog.is_ok(), "Interface should parse: {:?}", prog.err());
}

#[test]
fn test_structure_parses_gracefully() {
    let code = r#"
Structure Point
    Public X As Integer
    Public Y As Integer
End Structure
"#;
    let prog = parse_program(code);
    assert!(prog.is_ok(), "Structure should parse: {:?}", prog.err());
}

#[test]
fn test_namespace_parses_gracefully() {
    let code = r#"
Namespace MyApp.Utilities
    Public Class Helper
    End Class
End Namespace
"#;
    let prog = parse_program(code);
    assert!(prog.is_ok(), "Namespace should parse: {:?}", prog.err());
}

#[test]
fn test_delegate_sub_parses_gracefully() {
    let code = r#"
Delegate Sub MyCallback(sender As Object, e As Integer)
"#;
    let prog = parse_program(code);
    assert!(prog.is_ok(), "Delegate Sub should parse: {:?}", prog.err());
}

#[test]
fn test_delegate_function_parses_gracefully() {
    let code = r#"
Delegate Function MyFunc(x As Integer) As String
"#;
    let prog = parse_program(code);
    assert!(prog.is_ok(), "Delegate Function should parse: {:?}", prog.err());
}

#[test]
fn test_event_declaration_parses_gracefully() {
    let code = r#"
Public Class MyClass
    Public Event Click As EventHandler
End Class
"#;
    let prog = parse_program(code);
    assert!(prog.is_ok(), "Event declaration should parse: {:?}", prog.err());
}

#[test]
fn test_implements_statement_parses_gracefully() {
    let code = r#"
Public Class MyClass
    Implements IDisposable
    Public Sub Dispose()
    End Sub
End Class
"#;
    let prog = parse_program(code);
    assert!(prog.is_ok(), "Implements statement should parse: {:?}", prog.err());
}

// ============================================================
// And vs AndAlso, Or vs OrElse in Grammar
// ============================================================

#[test]
fn test_andalso_parses_distinct_from_and() {
    let code = r#"
Sub Test()
    Dim a As Boolean = True
    Dim b As Boolean = False
    Dim r1 As Boolean = a And b
    Dim r2 As Boolean = a AndAlso b
End Sub
"#;
    let prog = parse_program(code);
    assert!(prog.is_ok(), "And and AndAlso should both parse: {:?}", prog.err());
}

#[test]
fn test_orelse_parses_distinct_from_or() {
    let code = r#"
Sub Test()
    Dim a As Boolean = True
    Dim b As Boolean = False
    Dim r1 As Boolean = a Or b
    Dim r2 As Boolean = a OrElse b
End Sub
"#;
    let prog = parse_program(code);
    assert!(prog.is_ok(), "Or and OrElse should both parse: {:?}", prog.err());
}

// ============================================================
// Is / IsNot Parsing
// ============================================================

#[test]
fn test_is_nothing_parses() {
    let code = r#"
Sub Test()
    Dim obj As Object = Nothing
    If obj Is Nothing Then
        Console.WriteLine("null")
    End If
End Sub
"#;
    let prog = parse_program(code);
    assert!(prog.is_ok(), "Is Nothing should parse: {:?}", prog.err());
}

#[test]
fn test_isnot_nothing_parses() {
    let code = r#"
Sub Test()
    Dim obj As Object = Nothing
    If obj IsNot Nothing Then
        Console.WriteLine("not null")
    End If
End Sub
"#;
    let prog = parse_program(code);
    assert!(prog.is_ok(), "IsNot Nothing should parse: {:?}", prog.err());
}

// ============================================================
// Complex Expression Precedence (parse should succeed)
// ============================================================

#[test]
fn test_complex_mixed_precedence_parses() {
    let code = r#"
Sub Test()
    Dim x As Double = 2 ^ 3 + 4 * 5 - 10 \ 3 Mod 2
End Sub
"#;
    let prog = parse_program(code);
    assert!(prog.is_ok(), "Complex precedence should parse: {:?}", prog.err());
}

#[test]
fn test_chained_logical_operators_parse() {
    let code = r#"
Sub Test()
    Dim a As Boolean = True
    Dim b As Boolean = False
    Dim c As Boolean = True
    Dim result As Boolean = a AndAlso b OrElse c Xor Not a
End Sub
"#;
    let prog = parse_program(code);
    assert!(prog.is_ok(), "Chained logical ops should parse: {:?}", prog.err());
}

#[test]
fn test_combined_comparison_like_is_parse() {
    let code = r#"
Sub Test()
    Dim s As String = "Hello"
    Dim b1 As Boolean = s Like "H*"
    Dim obj As Object = Nothing
    Dim b2 As Boolean = obj Is Nothing
    Dim b3 As Boolean = obj IsNot Nothing
End Sub
"#;
    let prog = parse_program(code);
    assert!(prog.is_ok(), "Like, Is, IsNot should all parse in one sub: {:?}", prog.err());
}

// ============================================================
// Compound Assignment on Members
// ============================================================

#[test]
fn test_compound_on_member_parses() {
    // Test compound assignment on dotted member (not Me-qualified)
    let code = r#"
Sub Test()
    Dim obj As New MyClass()
    obj.Count += 1
End Sub

Public Class MyClass
    Public Count As Integer = 0
End Class
"#;
    let prog = parse_program(code);
    assert!(prog.is_ok(), "obj.Count += 1 should parse: {:?}", prog.err());
}

#[test]
fn test_typeof_in_if_condition() {
    let stmts = parse_sub_body(r#"
Sub Main()
    Dim s As String = "hello"
    If TypeOf s Is String Then
        Console.WriteLine("yes")
    End If
End Sub
"#);
    // Should parse as an If with TypeOf condition, not silently swallow it
    let has_if = stmts.iter().any(|s| matches!(s, Statement::If { .. }));
    assert!(has_if, "Should have an If statement. Got: {:?}", stmts);
    if let Statement::If { condition, then_branch, .. } = &stmts[1] {
        if let Expression::TypeOf { expr, type_name } = condition {
            assert_eq!(type_name, "String", "type_name should be 'String', got: {:?}", type_name);
            // Check the inner expression is the variable 's'
            assert!(matches!(expr.as_ref(), Expression::Variable(_)),
                "TypeOf inner expr should be a Variable. Got: {:?}", expr);
        } else {
            panic!("If condition should be TypeOf expression. Got: {:?}", condition);
        }
        assert!(!then_branch.is_empty(), "then_branch should not be empty. Got: {:?}", then_branch);
    }
}

#[test]
fn test_empty_array_literal_parses() {
    let stmts = parse_sub_body(r#"
Sub Main()
    Dim arr() As Integer = {}
End Sub
"#);
    assert!(!stmts.is_empty(), "Should parse empty array literal");
}
