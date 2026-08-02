' vybe-test: vb/vb_event_delegate_signature_matching/test_vb_delegate_equality_operator
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

Delegate Sub SimpleAction()

Module Program
    Private Sub Action1() : End Sub
    Private Sub Action2() : End Sub

    Sub Main()
        Dim d1 As SimpleAction = AddressOf Action1
        Dim d2 As SimpleAction = AddressOf Action1
        Dim d3 As SimpleAction = AddressOf Action2
        __Check(CStr((d1 = d2) & "|" & (d1 = d3)), "True|False")
    End Sub
End Module
