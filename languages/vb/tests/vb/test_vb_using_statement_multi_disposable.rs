use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Using Statement Multiple Variables Declaration
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_using_multiple_resources_comma_separated() {
    let src = r#"
Imports System

Class Tracker
    Implements IDisposable
    Public Name As String
    Public Sub New(n As String)
        Me.Name = n
    End Sub
    Public Sub Dispose() Implements IDisposable.Dispose
        Console.WriteLine("Disposed:" & Name)
    End Sub
End Class

Module Program
    Sub Main()
        Using r1 As New Tracker("R1"), r2 As New Tracker("R2")
            Console.WriteLine("Inside block")
        End Using
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["Inside block", "Disposed:R2", "Disposed:R1"]
    );
}
