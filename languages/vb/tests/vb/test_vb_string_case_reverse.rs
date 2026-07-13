use super::helpers::run_vb;

#[test]
fn string_case_reverse() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim s As String = "ViSuAl bAsIc"
        
        Console.WriteLine(UCase(s))
        Console.WriteLine(LCase(s))
        Console.WriteLine(StrReverse("stressed"))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["VISUAL BASIC", "visual basic", "desserts"]);
}
