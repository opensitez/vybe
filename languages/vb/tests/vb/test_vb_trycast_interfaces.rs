use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: TryCast on Interfaces
// ═══════════════════════════════════════════════════════════

#[test]
fn trycast_interfaces() {
    let out = run_vb(
        r#"
Interface ISpeak
    Sub Speak()
End Interface

Class Dog
    Implements ISpeak
    Public Sub Speak() Implements ISpeak.Speak
        Console.WriteLine("Woof")
    End Sub
End Class

Class Cat
    Public Sub Meow()
        Console.WriteLine("Meow")
    End Sub
End Class

Module M
    Sub Main()
        Dim d As Object = New Dog()
        Dim c As Object = New Cat()
        
        Dim s1 As ISpeak = TryCast(d, ISpeak)
        If s1 IsNot Nothing Then
            s1.Speak()
        End If
        
        Dim s2 As ISpeak = TryCast(c, ISpeak)
        If s2 Is Nothing Then
            Console.WriteLine("Cat does not implement ISpeak")
        End If
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Woof", "Cat does not implement ISpeak"]);
}
