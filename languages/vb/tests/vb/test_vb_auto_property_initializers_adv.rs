use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Auto-Implemented Property Initializers
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_auto_property_default_initializer() {
    let src = r#"
Class Config
    Public Property Port As Integer = 8080
    Public Property Host As String = "localhost"
    Public Property IsEnabled As Boolean = True
End Class

Module Program
    Sub Main()
        Dim cfg As New Config()
        Console.WriteLine(cfg.Host & ":" & cfg.Port & ":" & cfg.IsEnabled)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["localhost:8080:True"]);
}

#[test]
fn test_vb_auto_property_readonly_initializer() {
    let src = r#"
Class ImmutablePoint
    Public ReadOnly Property X As Double = 1.0
    Public ReadOnly Property Y As Double = 2.0
End Class

Module Program
    Sub Main()
        Dim pt As New ImmutablePoint()
        Console.WriteLine(pt.X & "," & pt.Y)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1,2"]);
}

#[test]
fn test_vb_auto_property_initializer_expression() {
    let src = r#"
Imports System.Collections.Generic

Class ShoppingCart
    Public Property Items As New List(Of String) From {"Item1", "Item2"}
End Class

Module Program
    Sub Main()
        Dim cart As New ShoppingCart()
        Console.WriteLine(cart.Items.Count)
        Console.WriteLine(cart.Items(0))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2", "Item1"]);
}

#[test]
fn test_vb_auto_property_override_in_constructor() {
    let src = r#"
Class Settings
    Public Property MaxRetries As Integer = 3

    Public Sub New(retries As Integer)
        Me.MaxRetries = retries
    End Sub
End Class

Module Program
    Sub Main()
        Dim s As New Settings(10)
        Console.WriteLine(s.MaxRetries)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10"]);
}
