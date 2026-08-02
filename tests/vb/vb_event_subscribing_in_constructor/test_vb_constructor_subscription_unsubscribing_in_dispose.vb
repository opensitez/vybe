' vybe-test: vb/vb_event_subscribing_in_constructor/test_vb_constructor_subscription_unsubscribing_in_dispose
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

Class ManagedSubscriber
    Implements IDisposable
    Private publisher As Publisher

    Public Sub New(p As Publisher)
        publisher = p
        AddHandler publisher.Notice, AddressOf OnNotice
    End Sub

    Private Sub OnNotice(sender As Object, e As EventArgs)
        __Check(CStr("Managed Notice"), "Managed Notice")
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        If publisher IsNot Nothing Then
            RemoveHandler publisher.Notice, AddressOf OnNotice
            publisher = Nothing
        End If
    End Sub
End Class

Class Publisher
    Public Event Notice As EventHandler
    Public Sub Fire()
        RaiseEvent Notice(Me, EventArgs.Empty)
    End Sub
End Class

Module Program
    Sub Main()
        Dim p As New Publisher()
        Using subObj As New ManagedSubscriber(p)
            p.Fire()
        End Using
        p.Fire() ' Should NOT output after dispose!
    End Sub
End Module
