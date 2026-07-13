use super::helpers::run_vb;

#[test]
fn mustinherit_mustoverride() {
    let out = run_vb(
        r#"
MustInherit Class Shape
    Public MustOverride Function GetArea() As Double
    
    Public Sub Print()
        Console.WriteLine("Area: " & GetArea())
    End Sub
End Class

Class Square
    Inherits Shape
    
    Public Property Side As Double
    
    Public Overrides Function GetArea() As Double
        Return Side * Side
    End Function
End Class

Module M
    Sub Main()
        Dim s As Shape = New Square() With {.Side = 5}
        s.Print()
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Area: 25"]);
}
