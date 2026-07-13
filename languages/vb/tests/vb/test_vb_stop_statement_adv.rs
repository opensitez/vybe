use super::helpers::run_vb;

#[test]
fn stop_statement_adv() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Stop breaks into the debugger
        If False Then
            Stop
        End If
        Console.WriteLine("Parsed Stop")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Parsed Stop"]);
}
