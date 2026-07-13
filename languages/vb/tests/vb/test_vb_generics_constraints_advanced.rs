use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Generics Constraints Advanced
// ═══════════════════════════════════════════════════════════

#[test]
fn generics_constraints_new() {
    let out = run_vb(
        r#"
' The New constraint requires the type argument to have a parameterless constructor
Class Factory(Of T As New)
    Public Function CreateInstance() As T
        Return New T()
    End Function
End Class

Class Item
    Public Property Name As String = "DefaultItem"
End Class

Module M
    Sub Main()
        Dim f As New Factory(Of Item)()
        Dim i As Item = f.CreateInstance()
        Console.WriteLine(i.Name)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["DefaultItem"]);
}

#[test]
fn generics_constraints_structure_and_class() {
    let out = run_vb(
        r#"
' Structure constraint requires value type (excluding Nullable)
' Class constraint requires reference type
Class HolderRef(Of T As Class)
    Public Property Value As T
End Class

Class HolderVal(Of T As Structure)
    Public Property Value As T
End Class

Class Item
End Class

Module M
    Sub Main()
        Dim hr As New HolderRef(Of Item)()
        Dim hv As New HolderVal(Of Integer)()
        
        hr.Value = Nothing
        hv.Value = 42
        
        Console.WriteLine(hr.Value Is Nothing)
        Console.WriteLine(hv.Value)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["True", "42"]);
}
