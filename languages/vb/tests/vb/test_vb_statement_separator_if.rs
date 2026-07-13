use super::helpers::run_vb;

#[test]
fn statement_separator_if() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim x = 10
        ' Single-line If with multiple statements separated by colons
        If x = 10 Then Console.Write("A") : Console.WriteLine("B") Else Console.WriteLine("C")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["AB"]);
}
