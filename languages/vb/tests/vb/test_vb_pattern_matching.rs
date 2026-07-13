use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Pattern Matching (Select Case TypeOf, Is)
// ═══════════════════════════════════════════════════════════

#[test]
fn pattern_matching_select_case_typeof() {
    let out = run_vb(
        r#"
Module M
    Sub PrintType(obj As Object)
        Select Case obj
            Case i As Integer
                Console.WriteLine("Integer: " & i.ToString())
            Case s As String
                Console.WriteLine("String: " & s)
            Case Else
                Console.WriteLine("Unknown")
        End Select
    End Sub

    Sub Main()
        PrintType(42)
        PrintType("Hello")
        PrintType(5.5)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Integer: 42", "String: Hello", "Unknown"]);
}

#[test]
fn pattern_matching_is_operator() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim obj As Object = "Pattern"
        
        If TypeOf obj Is String Then
            Dim s As String = DirectCast(obj, String)
            Console.WriteLine("Matched: " & s)
        End If
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Matched: Pattern"]);
}
