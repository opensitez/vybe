use super::helpers::run_vb;

#[test]
fn multiline_strings_adv() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Multi-line strings
        Dim s = "Line1
Line2"
        Console.WriteLine(s.Contains(Environment.NewLine) Or s.Contains(Chr(10)))
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["True"]);
}
