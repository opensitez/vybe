' vybe-test: vb/vb_on_error_resume_next/on_error_goto
' origin: languages/vb/tests/vb/test_vb_on_error_resume_next.rs

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
