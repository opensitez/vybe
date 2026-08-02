' vybe-test: vb/vb_null_reference_exception_guards/test_vb_null_conditional_operator_method_call
' origin: languages/vb/tests/vb/test_vb_null_reference_exception_guards.rs

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

Class Document
    Public Function GetTitle() As String
        Return "ValidTitle"
    End Function
End Class

Module Program
    Sub Main()
        Dim doc As Document = Nothing
        Dim title As String = doc?.GetTitle()
        __Check(CStr(title Is Nothing), "True")
    End Sub
End Module
