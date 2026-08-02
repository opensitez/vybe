' vybe-test: vb/vb_interop/b69_if_elseif_else
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

Function Classify(n As Integer) As String
    If n > 0 Then
        Return "positive"
    ElseIf n < 0 Then
        Return "negative"
    Else
        Return "zero"
    End If
End Function
__Check(CStr(Classify(5)), "positive")
__Check(CStr(Classify(-3)), "negative")
__Check(CStr(Classify(0)), "zero")
