use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Returning ValueTuples from Methods
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_method_returning_named_tuple() {
    let src = r#"
Module Program
    Function GetMinMax(numbers As Integer()) As (Min As Integer, Max As Integer)
        Dim minVal As Integer = numbers(0)
        Dim maxVal As Integer = numbers(0)
        For Each n In numbers
            If n < minVal Then minVal = n
            If n > maxVal Then maxVal = n
        Next
        Return (minVal, maxVal)
    End Function

    Sub Main()
        Dim res = GetMinMax({5, 2, 9, 1, 7})
        Console.WriteLine("Min=" & res.Min & ", Max=" & res.Max)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Min=1, Max=9"]);
}

#[test]
fn test_vb_try_pattern_with_tuple_return() {
    let src = r#"
Module Program
    Function TryParseInt(input As String) As (Success As Boolean, Value As Integer)
        Dim result As Integer
        If Integer.TryParse(input, result) Then
            Return (True, result)
        End If
        Return (False, 0)
    End Function

    Sub Main()
        Dim r1 = TryParseInt("123")
        Dim r2 = TryParseInt("abc")
        Console.WriteLine(r1.Success & ":" & r1.Value)
        Console.WriteLine(r2.Success & ":" & r2.Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True:123", "False:0"]);
}
