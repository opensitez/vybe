use super::helpers::run_vb;

#[test]
fn goto_string_label() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim i = 0
        
    StartLoop:
        If i = 2 Then GoTo EndLoop
        i += 1
        GoTo StartLoop
        
    EndLoop:
        Console.WriteLine(i)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["2"]);
}
