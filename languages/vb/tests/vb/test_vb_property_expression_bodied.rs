use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Expression-Bodied Properties & Members
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_property_single_line_getter() {
    let src = r#"
Class Person
    Public Property FirstName As String = "John"
    Public Property LastName As String = "Doe"

    Public ReadOnly Property FullName As String
        Get
            Return FirstName & " " & LastName
        End Get
    End Property
End Class

Module Program
    Sub Main()
        Dim p As New Person()
        Console.WriteLine(p.FullName)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["John Doe"]);
}

#[test]
fn test_vb_function_single_line_lambda_expression() {
    let src = r#"
Module Program
    Public Function Multiply(x As Integer, y As Integer) As Integer => x * y
    Public Function IsPositive(n As Integer) As Boolean => n > 0

    Sub Main()
        Console.WriteLine(Multiply(3, 4))
        Console.WriteLine(IsPositive(5))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["12", "True"]);
}

#[test]
fn test_vb_sub_single_line_expression() {
    let src = r#"
Module Program
    Public Sub LogMessage(msg As String) => Console.WriteLine("[LOG] " & msg)

    Sub Main()
        LogMessage("Test")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["[LOG] Test"]);
}
