' vybe-test: vb/vb_event_delegate_signature_matching/test_vb_delegate_multicast_invocation_list
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

Delegate Sub MultiNotify(msg As String)

Module Program
    Private Sub Logger1(msg As String) : __Check(CStr("Log1: " & msg), "Log1: Message") : End Sub
    Private Sub Logger2(msg As String) : __Check(CStr("Log2: " & msg), "Log2: Message") : End Sub

    Sub Main()
        Dim d As MultiNotify = AddressOf Logger1
        d = CType([Delegate].Combine(d, New MultiNotify(AddressOf Logger2)), MultiNotify)
        d("Message")
    End Sub
End Module
