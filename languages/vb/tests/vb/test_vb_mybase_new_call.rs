use super::helpers::run_vb;

#[test]
fn mybase_new_call() {
    let out = run_vb(
        r#"
Class Base
    Public Sub New(name As String)
        Console.WriteLine("Base: " & name)
    End Sub
End Class

Class Derived
    Inherits Base
    
    Public Sub New()
        ' Calling the base constructor explicitly using MyBase.New
        MyBase.New("Default")
        Console.WriteLine("Derived")
    End Sub
End Class

Module M
    Sub Main()
        Dim d As New Derived()
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Base: Default", "Derived"]);
}
