use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Overloads Statement
// ═══════════════════════════════════════════════════════════

#[test]
fn statement_overloads() {
    let out = run_vb(
        r#"
Class Base
    Public Sub Process(x As Integer)
        Console.WriteLine("Process Integer: " & x)
    End Sub
End Class

Class Derived
    Inherits Base
    
    ' Overloads is technically optional when the signatures are different,
    ' but it's used to explicitly define overloaded methods across inheritance bounds
    Public Overloads Sub Process(x As String)
        Console.WriteLine("Process String: " & x)
    End Sub
End Class

Module M
    Sub Main()
        Dim d As New Derived()
        d.Process(10)
        d.Process("Hello")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Process Integer: 10", "Process String: Hello"]);
}
