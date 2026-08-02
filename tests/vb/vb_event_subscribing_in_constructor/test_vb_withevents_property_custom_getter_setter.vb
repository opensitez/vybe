' vybe-test: vb/vb_event_subscribing_in_constructor/test_vb_withevents_property_custom_getter_setter
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

Class Notifier
    Public Event Ping As EventHandler
    Public Sub Fire()
        RaiseEvent Ping(Me, EventArgs.Empty)
    End Sub
End Class

Class ExplicitPropertyListener
    Private _notifier As Notifier

    Public Custom WithEvents Property NotifierProp As Notifier
        Get
            Return _notifier
        End Get
        Set(value As Notifier)
            If _notifier IsNot Nothing Then
                RemoveHandler _notifier.Ping, AddressOf OnPing
            End If
            _notifier = value
            If _notifier IsNot Nothing Then
                AddHandler _notifier.Ping, AddressOf OnPing
            End If
        End Set
    End Property

    Private Sub OnPing(sender As Object, e As EventArgs)
        __Check(CStr("Explicit Property Handled Ping"), "Explicit Property Handled Ping")
    End Sub
End Class

Module Program
    Sub Main()
        Dim n As New Notifier()
        Dim listener As New ExplicitPropertyListener()
        listener.NotifierProp = n
        n.Fire()
    End Sub
End Module
