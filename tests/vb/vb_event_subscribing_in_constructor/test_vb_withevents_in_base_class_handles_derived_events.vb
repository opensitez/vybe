' vybe-test: vb/vb_event_subscribing_in_constructor/test_vb_withevents_in_base_class_handles_derived_events
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

Class EventSource
    Public Event Alert As EventHandler
    Public Sub Trigger()
        RaiseEvent Alert(Me, EventArgs.Empty)
    End Sub
End Class

Class BaseListener
    Protected WithEvents Source As EventSource

    Public Sub New(s As EventSource)
        Source = s
    End Sub

    Private Sub OnAlert(sender As Object, e As EventArgs) Handles Source.Alert
        __Check(CStr("Base Listener Handled Alert"), "Base Listener Handled Alert")
    End Sub
End Class

Class DerivedListener
    Inherits BaseListener

    Public Sub New(s As EventSource)
        MyBase.New(s)
    End Sub
End Class

Module Program
    Sub Main()
        Dim s As New EventSource()
        Dim dl As New DerivedListener(s)
        s.Trigger()
    End Sub
End Module
