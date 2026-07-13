use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Classes (Shadows Keyword)
// ═══════════════════════════════════════════════════════════

#[test]
fn class_shadows_method() {
    let out = run_vb(
        r#"
Class Parent
    Public Sub ShowMessage()
        Console.WriteLine("Parent Message")
    End Sub
End Class

Class Child
    Inherits Parent
    
    Public Shadows Sub ShowMessage()
        Console.WriteLine("Child Message")
    End Sub
End Class

Module M
    Sub Main()
        Dim c As New Child()
        c.ShowMessage()
        
        Dim p As Parent = c
        p.ShowMessage() ' Should print Parent Message due to shadowing (not overriding)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Child Message", "Parent Message"]);
}

#[test]
fn class_shadows_field_with_property() {
    let out = run_vb(
        r#"
Class BaseData
    Public Value As String = "BaseValue"
End Class

Class DerivedData
    Inherits BaseData
    
    Private _val As String = "DerivedValue"
    Public Shadows Property Value As String
        Get
            Return _val
        End Get
        Set(v As String)
            _val = v
        End Set
    End Property
End Class

Module M
    Sub Main()
        Dim d As New DerivedData()
        Console.WriteLine(d.Value)
        
        Dim b As BaseData = d
        Console.WriteLine(b.Value)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["DerivedValue", "BaseValue"]);
}
