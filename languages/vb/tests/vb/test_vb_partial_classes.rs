use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Partial Classes
// ═══════════════════════════════════════════════════════════

#[test]
fn partial_class_fields() {
    let out = run_vb(
        r#"
Partial Class Employee
    Public FirstName As String
End Class

Partial Class Employee
    Public LastName As String
End Class

Module M
    Sub Main()
        Dim e As New Employee()
        e.FirstName = "John"
        e.LastName = "Smith"
        Console.WriteLine(e.FirstName & " " & e.LastName)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["John Smith"]);
}

#[test]
fn partial_class_methods() {
    let out = run_vb(
        r#"
Partial Class Calculator
    Public Function Add(a As Integer, b As Integer) As Integer
        Return a + b
    End Function
End Class

Partial Class Calculator
    Public Function Multiply(a As Integer, b As Integer) As Integer
        Return a * b
    End Function
End Class

Module M
    Sub Main()
        Dim c As New Calculator()
        Console.WriteLine(c.Add(2, 3))
        Console.WriteLine(c.Multiply(2, 3))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["5", "6"]);
}
