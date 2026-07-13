use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Late Binding (Properties)
// ═══════════════════════════════════════════════════════════

#[test]
fn late_binding_property_get_set() {
    let out = run_vb(
        r#"
Class DataModel
    Private _value As String
    Public Property Value As String
        Get
            Return _value
        End Get
        Set(v As String)
            _value = v
        End Set
    End Property
End Class

Module M
    Sub Main()
        Dim model As Object = New DataModel()
        
        ' Late bound property setter
        model.Value = "Dynamic Property"
        
        ' Late bound property getter
        Console.WriteLine(model.Value)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Dynamic Property"]);
}

#[test]
fn late_binding_indexed_property() {
    let out = run_vb(
        r#"
Class ItemCollection
    Private items As String() = {"A", "B", "C"}
    
    Default Public Property Item(index As Integer) As String
        Get
            Return items(index)
        End Get
        Set(value As String)
            items(index) = value
        End Set
    End Property
End Class

Module M
    Sub Main()
        Dim col As Object = New ItemCollection()
        
        ' Late bound indexed property via Default property (implicit)
        col(1) = "Z"
        
        ' Late bound indexed property (explicit)
        Console.WriteLine(col.Item(0))
        Console.WriteLine(col(1))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["A", "Z"]);
}
