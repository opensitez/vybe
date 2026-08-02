' vybe-test: vb/vb_return_in_catch_finally/return_in_catch_finally
' origin: languages/vb/tests/vb/test_vb_return_in_catch_finally.rs

' Vybe test harness — Visual Basic.
'
' Real VB source alongside harness/go/check.go and harness/js/check.js, the way
' test262's assert.js is JavaScript.
'
' A test's verdict is its EXIT CODE. __Check prints its diagnostic BEFORE
' throwing: an uncaught exception surfaces as `RuntimeError: [object]`, which
' says nothing at all.

Module VybeCheck
    Sub __Check(got As String, want As String)
        If got <> want Then
            Console.WriteLine("FAIL: want [" & want & "] got [" & got & "]")
            Throw New Exception("assertion failed")
        End If
    End Sub
End Module

Module M
    Function TestReturn() As Integer
        Try
            Throw New Exception("Error")
        Catch ex As Exception
            Return 1
        Finally
            ' VB.NET allows Return in Finally?
            ' Wait, Return in Finally is a compiler error in VB.NET (and C#)!
            ' So we just test Return in Catch and modifying the return value implicitly by assigning to the function name
            __Check(CStr("Finally executed"), "Finally executed")
        End Try
    End Function

    Function TestImplicitReturn() As Integer
        Try
            Throw New Exception("Error")
        Catch ex As Exception
            TestImplicitReturn = 2
            Exit Function
        Finally
            __Check(CStr("Finally executed 2"), "1")
        End Try
    End Function

    Sub Main()
        __Check(CStr(TestReturn()), "Finally executed 2")
        __Check(CStr(TestImplicitReturn()), "2")
    End Sub
End Module
