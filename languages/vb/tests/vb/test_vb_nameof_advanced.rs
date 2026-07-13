use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: NameOf Operator (Advanced)
// ═══════════════════════════════════════════════════════════

#[test]
fn nameof_operator_members() {
    let out = run_vb(
        r#"
Class Data
    Public Property Value As Integer
End Class

Module M
    Sub Main()
        ' NameOf can reference members of a type without an instance
        Console.WriteLine(NameOf(Data.Value))
        
        Dim d As New Data()
        Console.WriteLine(NameOf(d.Value))
        
        ' NameOf with local variables
        Dim localVariable As String = ""
        Console.WriteLine(NameOf(localVariable))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Value", "Value", "localVariable"]);
}
