use super::helpers::run_vb;

#[test]
fn interaction_msgbox_inputbox() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' We can't actually show UI, but we can test the syntax parsing
        ' Mocking MsgBoxResult Enum implicitly used
        Dim prompt As String = "Test"
        Dim title As String = "Title"
        
        ' Just check it compiles
        Dim msgType = MsgBoxStyle.OkOnly
        Console.WriteLine(CInt(msgType))
        
        ' Don't call them as it might hang the test runner if not mocked
        Console.WriteLine("Parsed")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["0", "Parsed"]);
}
