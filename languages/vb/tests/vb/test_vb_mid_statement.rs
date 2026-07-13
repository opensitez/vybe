use super::helpers::run_vb;

#[test]
fn mid_statement_mutation() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim text As String = "Hello World"
        
        ' Mid statement replaces part of a string (1-based index)
        Mid(text, 7, 5) = "VB.NET"
        
        ' It replaces up to the length specified or the end of the string
        ' "Hello World" (length 11). Replacing at 7 with 5 chars.
        ' "World" is replaced by "VB.NE"
        Console.WriteLine(text)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Hello VB.NE"]);
}

#[test]
fn mid_statement_no_length() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim text As String = "Apple"
        
        ' If length is omitted, it replaces as much as possible
        Mid(text, 2) = "nna"
        Console.WriteLine(text)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Annna"]);
}
