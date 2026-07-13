use super::helpers::run_vb;

#[test]
fn oop_polymorphism_shadowing() {
    let out = run_vb(
        r#"
Class Base
    Public Overridable Function GetName() As String
        Return "Base"
    End Function
End Class

Class Derived1
    Inherits Base
    Public Overrides Function GetName() As String
        Return "Derived1"
    End Function
End Class

Class Derived2
    Inherits Base
    ' Shadows the base method, doesn't override
    Public Shadows Function GetName() As String
        Return "Derived2"
    End Function
End Class

Module M
    Sub Main()
        Dim d1 As New Derived1()
        Dim d2 As New Derived2()
        
        Dim b1 As Base = d1
        Dim b2 As Base = d2
        
        Console.WriteLine(b1.GetName())
        Console.WriteLine(b2.GetName())
        Console.WriteLine(d2.GetName())
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Derived1", "Base", "Derived2"]);
}

#[test]
fn oop_mustinherit_mustoverride() {
    let out = run_vb(
        r#"
MustInherit Class Shape
    Public MustOverride Function Area() As Double
End Class

Class Circle
    Inherits Shape
    Public Radius As Double
    
    Public Overrides Function Area() As Double
        Return 3.14 * Radius * Radius
    End Function
End Class

Module M
    Sub Main()
        Dim c As New Circle() With {.Radius = 10}
        Dim s As Shape = c
        Console.WriteLine(s.Area())
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["314"]);
}
