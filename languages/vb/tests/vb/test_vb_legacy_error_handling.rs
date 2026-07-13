use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Legacy Error Handling (On Error GoTo, Resume Next)
// ═══════════════════════════════════════════════════════════

#[test]
fn on_error_goto_label() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        On Error GoTo ErrorHandler
        
        Dim a As Integer = 10
        Dim b As Integer = 0
        Dim c As Integer = a \ b
        
        Console.WriteLine("Should not print")
        Exit Sub

ErrorHandler:
        Console.WriteLine("Error caught")
        Console.WriteLine(Err.Number <> 0)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Error caught", "True"]);
}

#[test]
fn on_error_resume_next() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        On Error Resume Next
        
        Dim a As Integer = 10
        Dim b As Integer = 0
        Dim c As Integer = a \ b
        
        Console.WriteLine("Continuing execution")
        Console.WriteLine(Err.Number <> 0)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Continuing execution", "True"]);
}

#[test]
fn on_error_goto_0() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        On Error Resume Next
        Err.Raise(5, "Test", "Test Error")
        Console.WriteLine("Ignored")
        
        On Error GoTo 0 ' Disables error handling
        
        Try
            Err.Raise(6, "Test2", "Test Error 2")
        Catch ex As Exception
            Console.WriteLine("Caught by Try")
        End Try
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Ignored", "Caught by Try"]);
}

#[test]
fn err_object_properties() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        On Error GoTo Handler
        Err.Raise(1234, "MySource", "MyDescription")
        Exit Sub
        
Handler:
        Console.WriteLine(Err.Number)
        Console.WriteLine(Err.Source)
        Console.WriteLine(Err.Description)
        Err.Clear()
        Console.WriteLine(Err.Number)
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["1234", "MySource", "MyDescription", "0"]);
}

#[test]
fn resume_statement() {
    let out = run_vb(
        r#"
Module M
    Sub Main()
        Dim attempts As Integer = 0
        
        On Error GoTo Handler
RetryPoint:
        If attempts = 0 Then
            Dim x As Integer = 1 \ 0
        End If
        Console.WriteLine("Success")
        Exit Sub
        
Handler:
        attempts = attempts + 1
        Console.WriteLine("Attempt: " & attempts)
        Resume RetryPoint
    End Sub
End Module
"#,
    );
    assert_eq!(out, vec!["Attempt: 1", "Success"]);
}
