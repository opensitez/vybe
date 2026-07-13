use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Strings Advanced (Mid, InStr, InStrRev, StrReverse)
// ═══════════════════════════════════════════════════════════

#[test]
fn string_mid_function() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim text As String = "Hello Visual Basic"
        ' Mid is 1-based index in VB
        Console.WriteLine(Mid(text, 7, 6))
        ' Mid without length goes to the end
        Console.WriteLine(Mid(text, 14))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Visual", "Basic"]);
}

#[test]
fn string_mid_statement_assignment() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim text As String = "Hello Visual Basic"
        ' Mid can be used as a statement to replace parts of a string
        Mid(text, 7, 6) = "Modern"
        Console.WriteLine(text)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Hello Modern Basic"]);
}

#[test]
fn string_instr_function() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim text As String = "apple banana apple"
        ' InStr returns 1-based index
        Console.WriteLine(InStr(text, "banana"))
        ' Start search from index 8
        Console.WriteLine(InStr(8, text, "apple"))
        ' Case insensitive search (CompareMethod.Text = 1)
        Console.WriteLine(InStr(1, "Hello", "h", 1))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["7", "14", "1"]);
}

#[test]
fn string_instrrev_function() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim text As String = "apple banana apple"
        ' Searches from the right to left
        Console.WriteLine(InStrRev(text, "apple"))
        ' With start position
        Console.WriteLine(InStrRev(text, "apple", 10))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["14", "1"]);
}

#[test]
fn string_strreverse_function() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Console.WriteLine(StrReverse("stressed"))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["desserts"]);
}
