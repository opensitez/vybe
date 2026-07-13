use super::helpers::run_vb;

#[test]
fn myclass_keyword() {
    let out = run_vb(
        r#"
Class Base
    Public Overridable Sub Show()
        Console.WriteLine("Base")
    End Sub
    
    Public Sub CallShow()
        ' Me.Show() calls the overridden version in Derived
        Me.Show()
        ' MyClass.Show() statically calls the version defined in this class, bypassing overrides
        MyClass.Show()
    End Sub
End Class

Class Derived
    Inherits Base
    Public Overrides Sub Show()
        Console.WriteLine("Derived")
    End Sub
End Class

Module M
    Sub Main()
        Dim d As New Derived()
        d.CallShow()
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Derived", "Base"]);
}
