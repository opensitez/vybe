use super::helpers::run_vb;

#[test]
fn info_err_erl() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        On Error Resume Next
        
10:     Dim a = 1
20:     Error 5 ' Simulate an error on line 20
        
        ' Err object contains information about run-time errors
        Console.WriteLine(Err.Number)
        
        ' Erl function returns the line number where the error occurred
        Console.WriteLine(Erl())
        
        Err.Clear()
        Console.WriteLine(Err.Number)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["5", "20", "0"]);
}
