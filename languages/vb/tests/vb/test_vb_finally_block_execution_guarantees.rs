use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Finally Block Guarantees & Control Flow
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_finally_executes_on_return() {
    let src = r#"
Imports System

Module Program
    Function TestFunc() As String
        Try
            Return "ReturnedValue"
        Finally
            Console.WriteLine("FinallyExecuted")
        End Try
    End Function

    Sub Main()
        Dim res As String = TestFunc()
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["FinallyExecuted", "ReturnedValue"]);
}

#[test]
fn test_vb_finally_executes_on_uncaught_exception() {
    let src = r#"
Imports System

Module Program
    Sub SubWithFinally()
        Try
            Throw New InvalidOperationException("Fatal")
        Finally
            Console.WriteLine("Cleanup in Finally")
        End Try
    End Sub

    Sub Main()
        Try
            SubWithFinally()
        Catch ex As Exception
            Console.WriteLine("Caught in Main: " & ex.Message)
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["Cleanup in Finally", "Caught in Main: Fatal"]
    );
}
