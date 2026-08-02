' vybe-test: vb/vb_event_subscribing_in_constructor/test_vb_constructor_subscription_shared_static_event
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

Class GlobalNotifier
    Public Shared Event GlobalPing As EventHandler
    Public Shared Sub Fire()
        RaiseEvent GlobalPing(Nothing, EventArgs.Empty)
    End Sub
End Class

Class Subscriber
    Public Sub New()
        AddHandler GlobalNotifier.GlobalPing, AddressOf OnGlobalPing
    End Sub

    Private Sub OnGlobalPing(sender As Object, e As EventArgs)
        __Check(CStr("Global Ping Received"), "Global Ping Received")
    End Sub
End Class

Module Program
    Sub Main()
        Dim s As New Subscriber()
        GlobalNotifier.Fire()
    End Sub
End Module
