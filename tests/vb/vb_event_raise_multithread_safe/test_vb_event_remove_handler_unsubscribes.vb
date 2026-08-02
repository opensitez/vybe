' vybe-test: vb/vb_event_raise_multithread_safe/test_vb_event_remove_handler_unsubscribes
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

Class Alarm
    Public Event Ring As Action
    Public Sub Sound()
        RaiseEvent Ring()
    End Sub
End Class

Module Program
    Private Sub OnRing()
        __Check(CStr("Alarm Ringing"), "Alarm Ringing")
    End Sub

    Sub Main()
        Dim a As New Alarm()
        AddHandler a.Ring, AddressOf OnRing
        a.Sound()
        RemoveHandler a.Ring, AddressOf OnRing
        a.Sound()
    End Sub
End Module
