use super::helpers::run_vb;

#[test]
fn string_space_strdup() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Space returns a string of spaces
        Console.WriteLine("[" & Space(5) & "]")
        
        ' StrDup returns a string of repeating characters
        Console.WriteLine(StrDup(3, "A"))
        Console.WriteLine(StrDup(4, "*"c))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["[     ]", "AAA", "****"]);
}
