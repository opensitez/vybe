use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Partial Methods
// ═══════════════════════════════════════════════════════════

#[test]
fn partial_method_implementation() {
    let out = run_vb(
        r#"
Partial Class Logger
    ' Declaration
    Partial Private Sub LogMessage(msg As String)
    End Sub
    
    Public Sub DoWork()
        Console.WriteLine("Working")
        LogMessage("Work completed")
    End Sub
End Class

Partial Class Logger
    ' Implementation
    Private Sub LogMessage(msg As String)
        Console.WriteLine("LOG: " & msg)
    End Sub
End Class

Module M
    Sub Main()
        Dim l As New Logger()
        l.DoWork()
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Working", "LOG: Work completed"]);
}

#[test]
fn partial_method_no_implementation() {
    let out = run_vb(
        r#"
Partial Class SilencedLogger
    ' Declaration only, no implementation
    Partial Private Sub LogMessage(msg As String)
    End Sub
    
    Public Sub DoWork()
        Console.WriteLine("Working")
        ' If not implemented, call is removed entirely by compiler
        LogMessage("Work completed")
    End Sub
End Class

Module M
    Sub Main()
        Dim l As New SilencedLogger()
        l.DoWork()
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Working"]);
}
