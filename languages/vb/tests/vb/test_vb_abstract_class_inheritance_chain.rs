use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Abstract Class (MustInherit) & Inheritance Chain
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_mustinherit_class_hierarchy() {
    let src = r#"
MustInherit Class Shape
    Public MustOverride Function GetArea() As Double
End Class

MustInherit Class Polygon
    Inherits Shape
    Public MustOverride Function GetSides() As Integer
End Class

Class Rectangle
    Inherits Polygon
    Public Width As Double = 5.0
    Public Height As Double = 4.0

    Public Overrides Function GetArea() As Double
        Return Width * Height
    End Function

    Public Overrides Function GetSides() As Integer
        Return 4
    End Function
End Class

Module Program
    Sub Main()
        Dim rect As Shape = New Rectangle()
        Console.WriteLine(rect.GetArea())
        Dim poly As Polygon = CType(rect, Polygon)
        Console.WriteLine(poly.GetSides())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["20", "4"]);
}

#[test]
fn test_vb_mustinherit_concrete_base_methods() {
    let src = r#"
MustInherit Class BaseLogger
    Public Sub Log(msg As String)
        WriteEntry(FormatMessage(msg))
    End Sub

    Protected MustOverride Sub WriteEntry(formatted As String)

    Protected Virtual Function FormatMessage(msg As String) As String
        Return "[LOG] " & msg
    End Function
End Class

Class ConsoleLogger
    Inherits BaseLogger
    Protected Overrides Sub WriteEntry(formatted As String)
        Console.WriteLine(formatted)
    End Sub
End Class

Module Program
    Sub Main()
        Dim logger As BaseLogger = New ConsoleLogger()
        logger.Log("System initialized")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["[LOG] System initialized"]);
}

#[test]
fn test_vb_mustinherit_constructor_invocation() {
    let src = r#"
MustInherit Class Entity
    Public Property Id As Integer
    Protected Sub New(id As Integer)
        Me.Id = id
    End Sub
End Class

Class User
    Inherits Entity
    Public Property Name As String
    Public Sub New(id As Integer, name As String)
        MyBase.New(id)
        Me.Name = name
    End Sub
End Class

Module Program
    Sub Main()
        Dim u As New User(42, "Alice")
        Console.WriteLine(u.Id & ":" & u.Name)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["42:Alice"]);
}

#[test]
fn test_vb_mustinherit_abstract_property() {
    let src = r#"
MustInherit Class Component
    Public MustOverride Property Name As String
End Class

Class Button
    Inherits Component
    Private _name As String
    Public Overrides Property Name As String
        Get
            Return _name
        End Get
        Set(value As String)
            _name = value
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim btn As Component = New Button()
        btn.Name = "Submit"
        Console.WriteLine(btn.Name)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Submit"]);
}
