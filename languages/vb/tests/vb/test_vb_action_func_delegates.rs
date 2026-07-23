use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: System.Action & System.Func Generic Delegates
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_action_no_args() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim act As Action = Sub() Console.WriteLine("Action Ran")
        act()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Action Ran"]);
}

#[test]
fn test_vb_action_multiple_parameters() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim act As Action(Of String, Integer) = Sub(s, i) Console.WriteLine(s & ":" & i)
        act("Count", 5)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Count:5"]);
}

#[test]
fn test_vb_func_return_value() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim fn As Func(Of Integer, Integer, Integer) = Function(a, b) a * b
        Console.WriteLine(fn(6, 7))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["42"]);
}

#[test]
fn test_vb_predicate_delegate_matching() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim pred As Predicate(Of String) = Function(s) s.Length > 3
        Console.WriteLine(pred("Hi"))
        Console.WriteLine(pred("Hello"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False", "True"]);
}

#[test]
fn test_vb_func_chaining_higher_order_function() {
    let src = r#"
Imports System

Module Program
    Function Combine(f1 As Func(Of Integer, Integer), f2 As Func(Of Integer, Integer)) As Func(Of Integer, Integer)
        Return Function(x) f2(f1(x))
    End Function

    Sub Main()
        Dim addTwo As Func(Of Integer, Integer) = Function(x) x + 2
        Dim mulThree As Func(Of Integer, Integer) = Function(x) x * 3
        Dim combined As Func(Of Integer, Integer) = Combine(addTwo, mulThree)
        Console.WriteLine(combined(5))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["21"]);
}
