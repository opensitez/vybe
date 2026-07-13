use super::helpers::run_vb;

#[test]
fn multiple_statements_colon() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' The colon character allows placing multiple statements on a single line
        Dim x As Integer = 10 : Dim y As Integer = 20 : Console.WriteLine(x + y)
        
        If x = 10 Then : Console.WriteLine("Yes") : End If
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["30", "Yes"]);
}
