use super::helpers::run_vb;

#[test]
fn select_case_compare_text() {
    let out = run_vb(
        r#"
Option Compare Text

Module M
    Sub Main()
        Dim s = "hello"
        
        ' Select Case with Option Compare Text should be case insensitive
        Select Case s
            Case "HELLO"
                Console.WriteLine("Matched")
            Case Else
                Console.WriteLine("Not Matched")
        End Select
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Matched"]);
}
