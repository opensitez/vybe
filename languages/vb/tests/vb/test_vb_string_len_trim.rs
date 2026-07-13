use super::helpers::run_vb;

#[test]
fn string_len_trim() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim s As String = "  VB.NET  "
        
        ' Len measures string length (or variable byte size, but mostly string length)
        Console.WriteLine(Len(s))
        
        ' Trim functions
        Console.WriteLine("[" & Trim(s) & "]")
        Console.WriteLine("[" & LTrim(s) & "]")
        Console.WriteLine("[" & RTrim(s) & "]")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["10", "[VB.NET]", "[VB.NET  ]", "[  VB.NET]"]);
}
