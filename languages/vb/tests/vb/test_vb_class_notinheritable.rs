use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Classes (NotInheritable and NotOverridable)
// ═══════════════════════════════════════════════════════════

#[test]
fn class_notinheritable_basic() {
    let out = run_vb(
        r#"
NotInheritable Class MathUtils
    Public Function Add(a As Integer, b As Integer) As Integer
        Return a + b
    End Function
End Class

Module M
    Sub Main()
        Dim utils As New MathUtils()
        Console.WriteLine(utils.Add(10, 20))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["30"]);
}

#[test]
fn class_notoverridable_methods() {
    let out = run_vb(
        r#"
Class BasePrinter
    Public Overridable Sub Print()
        Console.WriteLine("Base")
    End Sub
End Class

Class FastPrinter
    Inherits BasePrinter
    
    ' Seals the method from further overriding in derived classes
    Public NotOverridable Overrides Sub Print()
        Console.WriteLine("Fast")
    End Sub
End Class

Module M
    Sub Main()
        Dim fp As BasePrinter = New FastPrinter()
        fp.Print()
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Fast"]);
}
