' vybe-test: vb/vb_event_delegate_signature_matching/test_vb_delegate_begin_invoke_end_invoke_async_pattern
' origin: languages/vb/tests/vb/test_vb_event_delegate_signature_matching.rs

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

Imports System

Delegate Function SlowCalc(n As Integer) As Integer

Module Program
    Private Function Compute(n As Integer) As Integer
        Return n * 10
    End Function

    Sub Main()
        Dim calc As SlowCalc = AddressOf Compute
        ' Invoke synchronously to test signature
        Dim res = calc.Invoke(5)
        __Check(CStr(res), "50")
    End Sub
End Module
