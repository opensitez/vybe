' vybe-test: vb/vb_event_delegate_signature_matching/test_vb_delegate_conversion_to_action
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

Module Program
    Private Sub Output(msg As String)
        __Check(CStr("ActionMsg: " & msg), "ActionMsg: Hello Action")
    End Sub

    Sub Main()
        Dim act As Action(Of String) = AddressOf Output
        act("Hello Action")
    End Sub
End Module
