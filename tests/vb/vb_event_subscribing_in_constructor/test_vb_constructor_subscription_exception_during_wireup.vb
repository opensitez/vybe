' vybe-test: vb/vb_event_subscribing_in_constructor/test_vb_constructor_subscription_exception_during_wireup
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

Class FaultyPublisher
    Public Custom Event CustomEvent As EventHandler
        AddHandler(value As EventHandler)
            Throw New InvalidOperationException("AddHandler Exception")
        End AddHandler
        RemoveHandler(value As EventHandler)
        End RemoveHandler
        RaiseEvent(sender As Object, e As EventArgs)
        End RaiseEvent
    End Event
End Class

Class FaultySubscriber
    Public Sub New(pub As FaultyPublisher)
        Try
            AddHandler pub.CustomEvent, Sub(s, e)
        Catch ex As InvalidOperationException
            __Check(CStr("Caught Exception During Wireup"), "Caught Exception During Wireup")
        End Try
    End Sub
End Class

Module Program
    Sub Main()
        Dim fp As New FaultyPublisher()
        Dim fs As New FaultySubscriber(fp)
    End Sub
End Module
