use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Generics (Classes)
// ═══════════════════════════════════════════════════════════

#[test]
fn generic_class_basic() {
    let out = run_vb(
        r#"
Class Box(Of T)
    Private _value As T
    
    Public Sub New(val As T)
        _value = val
    End Sub
    
    Public Function GetValue() As T
        Return _value
    End Function
End Class

Module M
    Sub Main()
        Dim intBox As New Box(Of Integer)(42)
        Dim strBox As New Box(Of String)("Hello")
        
        Console.WriteLine(intBox.GetValue())
        Console.WriteLine(strBox.GetValue())
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["42", "Hello"]);
}

#[test]
fn generic_class_multiple_type_parameters() {
    let out = run_vb(
        r#"
Class Pair(Of T1, T2)
    Public First As T1
    Public Second As T2
    
    Public Sub New(f As T1, s As T2)
        First = f
        Second = s
    End Sub
End Class

Module M
    Sub Main()
        Dim p As New Pair(Of String, Integer)("Age", 30)
        Console.WriteLine(p.First)
        Console.WriteLine(p.Second)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Age", "30"]);
}

#[test]
fn generic_class_inheritance() {
    let out = run_vb(
        r#"
Class Base(Of T)
    Public Value As T
End Class

Class DerivedInt
    Inherits Base(Of Integer)
End Class

Module M
    Sub Main()
        Dim d As New DerivedInt()
        d.Value = 99
        Console.WriteLine(d.Value)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["99"]);
}
