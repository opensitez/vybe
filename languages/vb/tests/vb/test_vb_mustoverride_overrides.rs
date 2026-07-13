use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: MustOverride and Overrides
// ═══════════════════════════════════════════════════════════

#[test]
fn mustoverride_and_overrides() {
    let out = run_vb(
        r#"
MustInherit Class Shape
    Public MustOverride Function GetArea() As Double
    Public MustOverride Property Name As String
End Class

Class Circle
    Inherits Shape
    
    Private _name As String = "Circle"
    Private _radius As Double
    
    Public Sub New(radius As Double)
        _radius = radius
    End Sub
    
    Public Overrides Function GetArea() As Double
        Return Math.PI * _radius * _radius
    End Function
    
    Public Overrides Property Name As String
        Get
            Return _name
        End Get
        Set(value As String)
            _name = value
        End Set
    End Property
End Class

Module M
    Sub Main()
        Dim s As Shape = New Circle(10)
        Console.WriteLine(s.Name)
        Console.WriteLine(Math.Round(s.GetArea()))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Circle", "314"]);
}
