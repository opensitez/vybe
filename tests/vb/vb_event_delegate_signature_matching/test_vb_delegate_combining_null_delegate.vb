' vybe-test: vb/vb_event_delegate_signature_matching/test_vb_delegate_combining_null_delegate
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

Delegate Sub SimpleDel()

Module Program
    Private Sub Target() : __Check(CStr("Target"), "Target") : End Sub

    Sub Main()
        Dim d1 As SimpleDel = Nothing
        Dim d2 As SimpleDel = AddressOf Target
        Dim combined As SimpleDel = CType([Delegate].Combine(d1, d2), SimpleDel)
        combined()
    End Sub
End Module
