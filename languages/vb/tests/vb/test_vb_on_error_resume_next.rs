use super::helpers::run_vb;

#[test]
fn on_error_resume_next_basic() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        ' Legacy error handling
        On Error Resume Next
        
        Dim x As Integer = 10
        Dim y As Integer = 0
        Dim z As Integer = x \ y ' Division by zero normally throws
        
        ' Execution continues to the next line
        Console.WriteLine("Continuing")
        Console.WriteLine(Err.Number <> 0) ' Err object holds the error
        
        Err.Clear()
        Console.WriteLine(Err.Number)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Continuing", "True", "0"]);
}

#[test]
fn on_error_goto() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        On Error GoTo ErrorHandler
        
        Throw New System.Exception("Test Error")
        Console.WriteLine("This won't print")
        Exit Sub
        
ErrorHandler:
        Console.WriteLine("Error Caught")
        Resume NextLine
        
NextLine:
        Console.WriteLine("Resumed")
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Error Caught", "Resumed"]);
}
