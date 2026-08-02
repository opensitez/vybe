' vybe-test: vb/vb_event_handler_add_remove_count/test_vb_custom_event_prevent_duplicate_handler
' origin: languages/vb/tests/vb/test_vb_event_handler_add_remove_count.rs

Imports System
Imports System.Collections.Generic

Class UniquePublisher
    Private handlerList As New List(Of EventHandler)()

    Public Custom Event UniqueEvent As EventHandler
        AddHandler(value As EventHandler)
            If Not handlerList.Contains(value) Then handlerList.Add(value)
        End AddHandler
        RemoveHandler(value As EventHandler)
            handlerList.Remove(value)
        End RemoveHandler
        RaiseEvent(sender As Object, e As EventArgs)
            For Each h In handlerList
                h(sender, e)
            Next
        End RaiseEvent
    End Event

    Public Sub Trigger()
        RaiseEvent UniqueEvent(Me, EventArgs.Empty)
    End Sub
End Class

Module Program
    Sub Main()
        Dim p As New UniquePublisher()
        Dim count = 0
        Dim handler As EventHandler = Sub(s, e) count += 1

        AddHandler p.UniqueEvent, handler
        AddHandler p.UniqueEvent, handler ' Duplicate add ignored by custom logic!
        p.Trigger()
        Console.WriteLine(count)
    End Sub
End Module
