' vybe-test: vb/vb_event_delegate_signature_matching/test_vb_delegate_target_property_instance_binding
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

Class TargetService
    Public Prefix As String
    Public Sub PrintMsg(msg As String)
        __Check(CStr(Prefix & ": " & msg), "True")
    End Sub
End Class

Delegate Sub MsgDelegate(msg As String)

Module Program
    Sub Main()
        Dim ts As New TargetService With {.Prefix = "LOG"}
        Dim d As MsgDelegate = AddressOf ts.PrintMsg
        __Check(CStr(d.Target Is ts), "LOG: Message text")
        d("Message text")
    End Sub
End Module
