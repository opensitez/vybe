use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Select Case Advanced (To, Is, Multiple expressions)
// ═══════════════════════════════════════════════════════════

#[test]
fn select_case_to_range() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim score As Integer = 85
        Select Case score
            Case 90 To 100
                Console.WriteLine("A")
            Case 80 To 89
                Console.WriteLine("B")
            Case 70 To 79
                Console.WriteLine("C")
            Case Else
                Console.WriteLine("F")
        End Select
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["B"]);
}

#[test]
fn select_case_is_operator() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim age As Integer = 15
        Select Case age
            Case Is >= 18
                Console.WriteLine("Adult")
            Case Is < 13
                Console.WriteLine("Child")
            Case Else
                Console.WriteLine("Teenager")
        End Select
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Teenager"]);
}

#[test]
fn select_case_multiple_expressions() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim dayOfWeek As Integer = 6
        Select Case dayOfWeek
            Case 1, 7
                Console.WriteLine("Weekend")
            Case 2, 3, 4, 5, 6
                Console.WriteLine("Weekday")
            Case Else
                Console.WriteLine("Invalid")
        End Select
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Weekday"]);
}

#[test]
fn select_case_mixed_conditions() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim value As Integer = 50
        Select Case value
            Case 1 To 10, 20 To 30, Is >= 100
                Console.WriteLine("Group 1")
            Case 40, 50, 60
                Console.WriteLine("Group 2")
            Case Else
                Console.WriteLine("Other")
        End Select
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Group 2"]);
}
