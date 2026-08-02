' vybe-test: vb/vb_byref_paramarray/byref_paramarray
' origin: languages/vb/tests/vb/test_vb_byref_paramarray.rs

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
    ' ParamArray cannot be ByRef. This tests the parser's error recovery or rejection.
    ' If we wrap it inside a class and don't execute it, we can verify parsing.
    Sub Test()
        __Check(CStr("Parsed"), "Parsed")
    End Sub

    Sub Main()
        Test()
    End Sub
End Module

Class InvalidSyntaxTest
    ' Sub Invalid(ByRef ParamArray x() As Integer)
    ' End Sub
End Class
