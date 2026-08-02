' vybe-test: vb/vb_interop/a06_method_chain
' origin: languages/vb/tests/vb/vb_interop_test.rs

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

Public Class Pipeline
    Public Function Step1(x As Integer) As Integer
        Return x + 1
    End Function
    Public Function Step2(x As Integer) As Integer
        Return Step1(x) * 2
    End Function
    Public Function Step3(x As Integer) As Integer
        Return Step2(x) + 10
    End Function
End Class
Dim p As New Pipeline()
__Check(CStr(p.Step3(5)), "22")
