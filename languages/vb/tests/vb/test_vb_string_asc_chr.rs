use super::helpers::run_vb;

#[test]
fn string_asc_chr() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Asc / AscW gets the integer code
        Console.WriteLine(Asc("A"))
        
        ' Chr / ChrW gets the char from integer
        Console.WriteLine(Chr(66))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["65", "B"]);
}
