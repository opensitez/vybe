use super::helpers::run_vb;

#[test]
fn appactivate_sendkeys() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim b As Boolean = True
        If Not b Then
            AppActivate("Calculator")
            SendKeys.SendWait("1{+}")
        End If
        Console.WriteLine("AppActivate Parsed")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["AppActivate Parsed"]);
}
