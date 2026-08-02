' vybe-test: vb/vb_event_delegate_signature_matching/test_vb_delegate_function_multicast_returns_last_result
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

Delegate Function Compute(x As Integer) As Integer

Module Program
    Private Function Square(x As Integer) As Integer : Return x * x : End Function
    Private Function Cube(x As Integer) As Integer : Return x * x * x : End Function

    Sub Main()
        Dim c As Compute = AddressOf Square
        c = CType([Delegate].Combine(c, New Compute(AddressOf Cube)), Compute)
        __Check(CStr(c(3)), "27")
    End Sub
End Module
