use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Classes (MustInherit and MustOverride)
// ═══════════════════════════════════════════════════════════

#[test]
fn class_mustinherit_basic() {
    let out = run_vb(
        r#"
MustInherit Class Animal
    Public Function Breathe() As String
        Return "Breathing"
    End Function
End Class

Class Cat
    Inherits Animal
End Class

Module M
    Sub Main()
        Dim c As New Cat()
        Console.WriteLine(c.Breathe())
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Breathing"]);
}

#[test]
fn class_mustoverride_methods() {
    let out = run_vb(
        r#"
MustInherit Class Shape
    Public MustOverride Function Area() As Integer
End Class

Class Square
    Inherits Shape
    
    Private _side As Integer
    Public Sub New(side As Integer)
        _side = side
    End Sub
    
    Public Overrides Function Area() As Integer
        Return _side * _side
    End Function
End Class

Module M
    Sub Main()
        Dim s As Shape = New Square(4)
        Console.WriteLine(s.Area())
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["16"]);
}

#[test]
fn class_mustoverride_properties() {
    let out = run_vb(
        r#"
MustInherit Class Vehicle
    Public MustOverride ReadOnly Property Wheels() As Integer
End Class

Class Tricycle
    Inherits Vehicle
    
    Public Overrides ReadOnly Property Wheels() As Integer
        Get
            Return 3
        End Get
    End Property
End Class

Module M
    Sub Main()
        Dim t As Vehicle = New Tricycle()
        Console.WriteLine(t.Wheels)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["3"]);
}
