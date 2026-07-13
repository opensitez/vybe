use super::helpers::run_vb;

#[test]
fn singleline_if_multiple() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim x = 10
        Dim y = 0
        Dim z = 0
        
        ' Single line If with multiple statements separated by colon
        If x = 10 Then y = 1 : z = 2
        
        Console.WriteLine(y & "-" & z)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["1-2"]);
}
