use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Advanced Casts (TryCast, DirectCast)
// ═══════════════════════════════════════════════════════════

#[test]
fn casts_trycast() {
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
        Dim a1 As Animal = New Dog()
        Dim a2 As Animal = New Animal()
        
        ' TryCast returns Nothing if the cast fails (only for reference types)
        Dim d1 As Dog = TryCast(a1, Dog)
        Dim d2 As Dog = TryCast(a2, Dog)
        
        If d1 IsNot Nothing Then d1.Bark()
        Console.WriteLine(d2 Is Nothing)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Woof", "True"]);
}

#[test]
fn casts_directcast() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim obj As Object = "Hello"
        
        ' DirectCast requires exact type match or inheritance (stricter than CType)
        Dim str As String = DirectCast(obj, String)
        Console.WriteLine(str)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Hello"]);
}
