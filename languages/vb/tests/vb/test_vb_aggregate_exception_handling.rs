use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: AggregateException & InnerExceptions Unwrapping
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_aggregate_exception_flatten_and_handle() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim ex1 As New InvalidOperationException("Op1 failed")
        Dim ex2 As New ArgumentException("Arg2 invalid")
        Dim agg As New AggregateException(ex1, ex2)

        Console.WriteLine(agg.InnerExceptions.Count)

        agg.Handle(Function(e)
            Console.WriteLine("Handled: " & e.GetType().Name)
            Return True
        End Function)
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec![
            "2",
            "Handled: InvalidOperationException",
            "Handled: ArgumentException"
        ]
    );
}

#[test]
fn test_vb_aggregate_exception_nested_flatten() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim innerAgg As New AggregateException(New Exception("E1"))
        Dim outerAgg As New AggregateException(innerAgg, New Exception("E2"))

        Dim flat As AggregateException = outerAgg.Flatten()
        Console.WriteLine(flat.InnerExceptions.Count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2"]);
}
