use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Structs (Value Type Semantics)
// ═══════════════════════════════════════════════════════════

#[test]
fn struct_value_type_copy() {
    let out = run_vb(
        r#"
Structure Point
    Public X As Integer
    Public Y As Integer
End Structure

Module M
    Sub Main()
        Dim p1 As Point
        p1.X = 10
        p1.Y = 20
        
        ' Assignment creates a copy
        Dim p2 As Point = p1
        p2.X = 50
        
        Console.WriteLine(p1.X)
        Console.WriteLine(p2.X)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["10", "50"]);
}

#[test]
fn struct_with_block() {
    let out = run_vb(
        r#"
Structure Rect
    Public Width As Integer
    Public Height As Integer
End Structure

Module M
    Sub Main()
        Dim r As Rect
        With r
            .Width = 100
            .Height = 200
        End With
        Console.WriteLine(r.Width)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["100"]);
}
