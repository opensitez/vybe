use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Using Statement Advanced
// ═══════════════════════════════════════════════════════════

#[test]
fn using_statement_multiple_resources() {
    let out = run_vb(
        r#"
Imports System

Class Resource
    Implements IDisposable
    
    Public Name As String
    
    Public Sub New(n As String)
        Name = n
        Console.WriteLine("Acquired " & Name)
    End Sub
    
    Public Sub Dispose() Implements IDisposable.Dispose
        Console.WriteLine("Disposed " & Name)
    End Sub
End Class

Module M
    Sub Main()
        ' Multiple resources of the same type can be declared in one Using block
        Using r1 As New Resource("R1"), r2 As New Resource("R2")
            Console.WriteLine("Using " & r1.Name & " and " & r2.Name)
        End Using
    End Sub
End Module
"#,
    );
    // Note: Disposal order is reverse of acquisition order in C# and VB.NET
    assert_eq!(
        out,
        vec![
            "Acquired R1",
            "Acquired R2",
            "Using R1 and R2",
            "Disposed R2",
            "Disposed R1"
        ]
    );
}
