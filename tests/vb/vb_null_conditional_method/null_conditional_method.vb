' vybe-test: vb/vb_null_conditional_method/null_conditional_method
' origin: languages/vb/tests/vb/test_vb_null_conditional_method.rs

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

Class Person
    Public Sub DoWork()
        __Check(CStr("Working"), "Working")
    End Sub
End Class

Module M
    Sub Main()
        Dim p As Person = Nothing
        
        ' Null conditional method call
        p?.DoWork()
        
        p = New Person()
        p?.DoWork()
    End Sub
End Module
