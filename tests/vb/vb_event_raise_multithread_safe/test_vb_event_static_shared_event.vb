' vybe-test: vb/vb_event_raise_multithread_safe/test_vb_event_static_shared_event
' origin: languages/vb/tests/vb/test_vb_event_raise_multithread_safe.rs

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

Class GlobalBus
    Public Shared Event GlobalMessage As Action(Of String)
    Public Shared Sub Broadcast(msg As String)
        RaiseEvent GlobalMessage(msg)
    End Sub
End Class

Module Program
    Sub Main()
        AddHandler GlobalBus.GlobalMessage, Sub(m) __Check(CStr("Global: " & m), "Global: Ping")
        GlobalBus.Broadcast("Ping")
    End Sub
End Module
