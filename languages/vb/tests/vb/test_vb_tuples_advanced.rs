use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Tuples Advanced (Deconstruction, Returning)
// ═══════════════════════════════════════════════════════════

#[test]
fn tuples_advanced_return_and_deconstruct() {
    let out = run_vb(
        r#"
Module M
    Function GetCoordinates() As (X As Integer, Y As Integer)
        Return (10, 20)
    End Function

    Sub Main()
        ' Return named tuple
        Dim coords = GetCoordinates()
        Console.WriteLine(coords.X)
        Console.WriteLine(coords.Y)
        
        ' Deconstruction into existing variables (or new ones)
        Dim a, b As Integer
        (a, b) = GetCoordinates()
        Console.WriteLine(a)
        Console.WriteLine(b)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["10", "20", "10", "20"]);
}
