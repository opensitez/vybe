' vybe-test: vb/vb_class_static_shared_constructor/test_vb_shared_constructor_event_subscription
' origin: languages/vb/tests/vb/test_vb_class_static_shared_constructor.rs

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

Class SystemMonitor
    Public Shared Event Heartbeat As EventHandler
    Shared Sub New()
        ' Subscribe internal logger
        AddHandler Heartbeat, Sub(sender, args) __Check(CStr("Internal Heartbeat Logged"), "Internal Heartbeat Logged")
    End Sub
    Public Shared Sub Pulse()
        RaiseEvent Heartbeat(Nothing, EventArgs.Empty)
    End Sub
End Class

Module Program
    Sub Main()
        SystemMonitor.Pulse()
    End Sub
End Module
