use super::helpers::run_vb;

#[test]
fn mybase_property() {
    let out = run_vb(
        r#"
Class Base
    Public Overridable ReadOnly Property Name As String
        Get
            Return "Base"
        End Get
    End Property
End Class

Class Derived
    Inherits Base
    
    Public Overrides ReadOnly Property Name As String
        Get
            Return MyBase.Name & "Derived"
        End Get
    End Property
End Class

Module M
    Sub Main()
        Dim d As New Derived()
        Console.WriteLine(d.Name)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["BaseDerived"]);
}
