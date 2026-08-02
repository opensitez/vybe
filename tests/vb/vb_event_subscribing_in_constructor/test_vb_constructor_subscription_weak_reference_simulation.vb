' vybe-test: vb/vb_event_subscribing_in_constructor/test_vb_constructor_subscription_weak_reference_simulation
' origin: languages/vb/tests/vb/test_vb_event_subscribing_in_constructor.rs

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

Class ShortLivedSubscriber
    Public Sub New(pub As Broadcaster)
        AddHandler pub.Broadcast, AddressOf HandleBroadcast
    End Sub

    Private Sub HandleBroadcast(sender As Object, e As EventArgs)
        __Check(CStr("ShortLived Handled"), "ShortLived Handled")
    End Sub
End Class

Class Broadcaster
    Public Event Broadcast As EventHandler
    Public Sub Fire()
        RaiseEvent Broadcast(Me, EventArgs.Empty)
    End Sub
End Class

Module Program
    Sub Main()
        Dim b As New Broadcaster()
        Dim subObj As New ShortLivedSubscriber(b)
        b.Fire()
    End Sub
End Module
