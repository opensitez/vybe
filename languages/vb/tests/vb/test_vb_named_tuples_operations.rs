use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Named ValueTuples Access & Compatibility
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_named_tuple_literal_access() {
    let src = r#"
Module Program
    Sub Main()
        Dim item As (Id As Integer, Name As String) = (1, "Widget")
        Console.WriteLine(item.Id & ":" & item.Name)
        Console.WriteLine(item.Item1 & ":" & item.Item2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1:Widget", "1:Widget"]);
}

#[test]
fn test_vb_named_tuple_inference_from_variable_names() {
    let src = r#"
Module Program
    Sub Main()
        Dim count As Integer = 5
        Dim label As String = "Total"
        Dim t = (count, label)
        Console.WriteLine(t.count & ":" & t.label)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["5:Total"]);
}

#[test]
fn test_vb_named_tuple_assignment_type_erasure() {
    let src = r#"
Module Program
    Sub Main()
        Dim t1 As (A As Integer, B As String) = (10, "Ten")
        Dim t2 As (X As Integer, Y As String) = t1 ' Type name erase assignment
        Console.WriteLine(t2.X & ":" & t2.Y)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10:Ten"]);
}
