' vybe-test: vb/vb_error_handling_try_catch/try_catch_when_clause_false
' origin: languages/vb/tests/vb/test_vb_error_handling_try_catch.rs

Module M
Sub Main()
Dim flag = False
Try
Throw New System.Exception()
Catch ex As System.Exception When flag
Console.WriteLine("Caught")
Catch
Console.WriteLine("Other")
End Try
End Sub
End Module
