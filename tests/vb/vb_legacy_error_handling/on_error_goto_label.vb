' vybe-test: vb/vb_legacy_error_handling/on_error_goto_label
' origin: languages/vb/tests/vb/test_vb_legacy_error_handling.rs

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
