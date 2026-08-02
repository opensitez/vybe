' vybe-test: vb/vb_callbyname_function_invocation/test_vb_callbyname_byref_argument_mutation
' origin: languages/vb/tests/vb/test_vb_callbyname_function_invocation.rs

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

Imports Microsoft.VisualBasic

Module Program
    Class Transformer
        Public Sub DoubleValue(ByRef x As Integer)
            x *= 2
        End Sub
    End Class

    Sub Main()
        Dim t As New Transformer()
        Dim val As Integer = 25
        CallByName(t, "DoubleValue", CallType.Method, val)
        __Check(CStr(val), "50")
    End Sub
End Module
