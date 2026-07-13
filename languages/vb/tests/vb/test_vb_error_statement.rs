use super::helpers::run_vb;

#[test]
fn error_statement() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        On Error GoTo ErrorHandler
        
        ' The Error statement simulates a runtime error
        Error 11 ' Division by zero error code
        Console.WriteLine("Unreachable")
        Exit Sub
        
ErrorHandler:
        Console.WriteLine("Error " & Err.Number)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Error 11"]);
}
