' vybe-test: vb/vb_custom_event_thread_sync/test_vb_custom_event_raise_with_custom_sender
' origin: languages/vb/tests/vb/test_vb_custom_event_thread_sync.rs

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

Class VirtualSender
End Class

Class CustomSenderBroadcaster
    Public Event Notice As EventHandler
    Public Sub FireWithSender(customSender As Object)
        RaiseEvent Notice(customSender, EventArgs.Empty)
    End Sub
End Class

Module Program
    Sub Main()
        Dim csb As New CustomSenderBroadcaster()
        Dim virt As New VirtualSender()
        AddHandler csb.Notice, Sub(s, e) __Check(CStr("Sender Type: " & s.GetType().Name), "Sender Type: VirtualSender")
        csb.FireWithSender(virt)
    End Sub
End Module
