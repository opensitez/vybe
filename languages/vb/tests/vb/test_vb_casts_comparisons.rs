use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Cast Comparisons (CType, DirectCast, TryCast)
// ═══════════════════════════════════════════════════════════

#[test]
fn casts_comparisons() {
    let out = run_vb(
        r#"
Class Animal
End Class

Class Dog
    Inherits Animal
    Public Sub Bark()
        Console.WriteLine("Woof")
    End Sub
End Class

Module M
    Sub Main()
        Dim a As Animal = New Dog()
        
        ' DirectCast requires an inheritance or implementation relationship
        Dim d1 As Dog = DirectCast(a, Dog)
        d1.Bark()
        
        ' TryCast returns Nothing if the cast fails (only for reference types)
        Dim a2 As New Animal()
        Dim d2 As Dog = TryCast(a2, Dog)
        If d2 Is Nothing Then
            Console.WriteLine("Cast Failed")
        End If
        
        ' CType can do conversions as well as casts (e.g. String to Integer)
        Dim numStr As Object = "123"
        Dim num As Integer = CType(numStr, Integer)
        Console.WriteLine(num + 1)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Woof", "Cast Failed", "124"]);
}
