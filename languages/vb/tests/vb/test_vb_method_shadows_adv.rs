use super::helpers::run_vb;

#[test]
fn method_shadows_adv() {
    let out = run_vb(
        r#"
Class Base
    Public Sub Process()
        Console.WriteLine("Base")
    End Sub
End Class

Class Derived
    Inherits Base
    
    ' Shadows hides the base method instead of overriding
    Public Shadows Sub Process()
        Console.WriteLine("Derived")
    End Sub
End Class

Module M
    Sub Main()
        Dim d As New Derived()
        Dim b As Base = d
        
        d.Process()
        b.Process()
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Derived", "Base"]);
}
