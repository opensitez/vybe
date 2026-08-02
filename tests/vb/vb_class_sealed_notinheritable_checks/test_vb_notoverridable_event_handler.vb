' vybe-test: vb/vb_class_sealed_notinheritable_checks/test_vb_notoverridable_event_handler
' origin: languages/vb/tests/vb/test_vb_class_sealed_notinheritable_checks.rs

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

Class BaseEmitter
    Public Overridable Custom Event Action As EventHandler
        AddHandler(value As EventHandler)
        End AddHandler
        RemoveHandler(value As EventHandler)
        End RemoveHandler
        RaiseEvent(sender As Object, e As EventArgs)
        End RaiseEvent
    End Event
End Class

Class FixedEmitter
    Inherits BaseEmitter
    Public NotOverridable Overrides Custom Event Action As EventHandler
        AddHandler(value As EventHandler)
            __Check(CStr("Handler Added to Fixed"), "Handler Added to Fixed")
        End AddHandler
        RemoveHandler(value As EventHandler)
        End RemoveHandler
        RaiseEvent(sender As Object, e As EventArgs)
        End RaiseEvent
    End Event
End Class

Module Program
    Sub Main()
        Dim e As BaseEmitter = New FixedEmitter()
        AddHandler e.Action, Sub(sender, args) End Sub
    End Sub
End Module
