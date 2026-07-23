use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: ValueTuple Deconstruction & Assignment
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_tuple_deconstruct_into_existing_vars() {
    let src = r#"
Module Program
    Sub Main()
        Dim t = (10, "Ten")
        Dim x As Integer
        Dim y As String
        (x, y) = t
        Console.WriteLine(x & ":" & y)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10:Ten"]);
}

#[test]
fn test_vb_tuple_deconstruct_custom_class() {
    let src = r#"
Class Point2D
    Public X As Double
    Public Y As Double

    Public Sub New(x As Double, y As Double)
        Me.X = x
        Me.Y = y
    End Sub

    Public Sub Deconstruct(ByRef outX As Double, ByRef outY As Double)
        outX = Me.X
        outY = Me.Y
    End Sub
End Class

Module Program
    Sub Main()
        Dim pt As New Point2D(3.0, 4.0)
        Dim px, py As Double
        (px, py) = pt
        Console.WriteLine(px & "," & py)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3,4"]);
}
