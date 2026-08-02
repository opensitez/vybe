' vybe-test: vb/vb_event_raise_multithread_safe/test_vb_event_custom_event_add_remove_raise_blocks
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

Class CustomPublisher
    Private handlerList As Action(Of String)

    Public Custom Event Message As Action(Of String)
        AddHandler(value As Action(Of String))
            handlerList = CType([Delegate].Combine(handlerList, value), Action(Of String))
            __Check(CStr("Custom AddHandler"), "Custom AddHandler")
        End AddHandler
        RemoveHandler(value As Action(Of String))
            handlerList = CType([Delegate].Remove(handlerList, value), Action(Of String))
            __Check(CStr("Custom RemoveHandler"), "Got: Hello")
        End RemoveHandler
        RaiseEvent(msg As String)
            If handlerList IsNot Nothing Then handlerList(msg)
        End RaiseEvent
    End Event

    Public Sub Dispatch(m As String)
        RaiseEvent Message(m)
    End Sub
End Class

Module Program
    Sub Main()
        Dim p As New CustomPublisher()
        Dim h As Action(Of String) = Sub(m) __Check(CStr("Got: " & m), "Custom RemoveHandler")
        AddHandler p.Message, h
        p.Dispatch("Hello")
        RemoveHandler p.Message, h
    End Sub
End Module
