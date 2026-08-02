' vybe-test: vb/vb_event_handler_add_remove_count/test_vb_custom_event_addhandler_removehandler_accessor
' origin: languages/vb/tests/vb/test_vb_event_handler_add_remove_count.rs

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

Class Button
    Private handlers As EventHandler

    Public Custom Event Click As EventHandler
        AddHandler(value As EventHandler)
            handlers = CType(Delegate.Combine(handlers, value), EventHandler)
        End AddHandler

        RemoveHandler(value As EventHandler)
            handlers = CType(Delegate.Remove(handlers, value), EventHandler)
        End RemoveHandler

        RaiseEvent(sender As Object, e As EventArgs)
            If handlers IsNot Nothing Then handlers(sender, e)
        End RaiseEvent
    End Event

    Public Sub PerformClick()
        RaiseEvent Click(Me, EventArgs.Empty)
    End Sub
End Class

Module Program
    Sub Main()
        Dim btn As New Button()
        Dim clicked = False
        Dim handler As EventHandler = Sub(s, e) clicked = True

        AddHandler btn.Click, handler
        btn.PerformClick()
        __Check(CStr(clicked), "True")

        clicked = False
        RemoveHandler btn.Click, handler
        btn.PerformClick()
        __Check(CStr(clicked), "False")
    End Sub
End Module
