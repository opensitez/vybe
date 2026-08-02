' vybe-test: vb/vb_custom_event_add_remove_raise/test_vb_custom_event_block_explicit_accessors
' origin: languages/vb/tests/vb/test_vb_custom_event_add_remove_raise.rs

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

Public Delegate Sub CustomHandler(msg As String)

Class Publisher
    Private _handlers As CustomHandler

    Public Custom Event StateChanged As CustomHandler
        AddHandler(value As CustomHandler)
            __Check(CStr("Added"), "Added")
            _handlers = CType([Delegate].Combine(_handlers, value), CustomHandler)
        End AddHandler

        RemoveHandler(value As CustomHandler)
            __Check(CStr("Removed"), "Raising")
            _handlers = CType([Delegate].Remove(_handlers, value), CustomHandler)
        End RemoveHandler

        RaiseEvent(msg As String)
            __Check(CStr("Raising"), "Received: Data1")
            _handlers?.Invoke(msg)
        End RaiseEvent
    End Event

    Public Sub Trigger(msg As String)
        RaiseEvent StateChanged(msg)
    End Sub
End Class

Module Program
    Sub Main()
        Dim pub As New Publisher()
        Dim h As CustomHandler = Sub(m) __Check(CStr("Received: " & m), "Removed")
        AddHandler pub.StateChanged, h
        pub.Trigger("Data1")
        RemoveHandler pub.StateChanged, h
    End Sub
End Module
