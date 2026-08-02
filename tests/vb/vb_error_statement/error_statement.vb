' vybe-test: vb/vb_error_statement/error_statement
' origin: languages/vb/tests/vb/test_vb_error_statement.rs

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
