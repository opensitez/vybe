use super::helpers::run_vb;

#[test]
fn declare_auto_function() {
    let out = run_vb(
        r#"
Module M
    ' Declare statement is used to call external DLLs
    ' We just test the syntax parsing here since we can't guarantee User32 is available in test environment
    Declare Auto Function MessageBox Lib "user32.dll" (ByVal hWnd As Integer, ByVal txt As String, ByVal caption As String, ByVal Typ As Integer) As Integer
    
    Sub Main()
        Dim parsed As Boolean = True
        Console.WriteLine(parsed)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["True"]);
}
