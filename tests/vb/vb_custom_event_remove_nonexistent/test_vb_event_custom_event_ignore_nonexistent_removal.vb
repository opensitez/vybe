' vybe-test: vb/vb_custom_event_remove_nonexistent/test_vb_event_custom_event_ignore_nonexistent_removal
' origin: languages/vb/tests/vb/test_vb_custom_event_remove_nonexistent.rs

Imports System

Class CustomManager
    Private handlers As Action

    Public Custom Event Work As Action
        AddHandler(value As Action)
            handlers = CType([Delegate].Combine(handlers, value), Action)
        End AddHandler
        RemoveHandler(value As Action)
            Dim newHandlers = CType([Delegate].Remove(handlers, value), Action)
            If newHandlers Is Nothing AndAlso handlers IsNot Nothing Then
                Console.WriteLine("All Handlers Removed")
            End If
            handlers = newHandlers
        End RemoveHandler
        RaiseEvent()
            If handlers IsNot Nothing Then handlers()
        End RaiseEvent
    End Event

    Public Sub Run()
        RaiseEvent Work()
    End Sub
End Class

Module Program
    Private Sub Sub1() : Console.WriteLine("Sub1") : End Sub

    Sub Main()
        Dim cm As New CustomManager()
        AddHandler cm.Work, AddressOf Sub1
        RemoveHandler cm.Work, AddressOf Sub1
    End Sub
End Module
