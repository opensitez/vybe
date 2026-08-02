' vybe-test: vb/vb_event_subscribing_in_constructor/test_vb_event_subscribed_in_constructor
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

Class InternalHandler
    Public Property Handled As Boolean = False
    Public Sub New(publisher As EventPublisher)
        AddHandler publisher.Triggered, AddressOf OnPublisherTriggered
    End Sub

    Private Sub OnPublisherTriggered(sender As Object, e As EventArgs)
        Handled = True
    End Sub
End Class

Class EventPublisher
    Public Event Triggered As EventHandler
    Public Sub Fire()
        RaiseEvent Triggered(Me, EventArgs.Empty)
    End Sub
End Class

Module Program
    Sub Main()
        Dim pub As New EventPublisher()
        Dim subObj As New InternalHandler(pub)
        pub.Fire()
        __Check(CStr(subObj.Handled), "True")
    End Sub
End Module
