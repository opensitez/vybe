use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Yield Return (State Machine / Logic)
// ═══════════════════════════════════════════════════════════

#[test]
fn yield_return_state_machine() {
    let out = run_vb(
        r#"
Imports System.Collections.Generic

Module M
    Iterator Function GetEvenNumbers(max As Integer) As IEnumerable(Of Integer)
        For i As Integer = 1 To max
            If i Mod 2 = 0 Then
                Yield i
            End If
        Next
    End Function

    Sub Main()
        For Each n In GetEvenNumbers(5)
            Console.WriteLine(n)
        Next
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["2", "4"]);
}
