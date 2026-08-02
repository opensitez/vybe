' vybe-test: vb/vb_custom_event_accessors/custom_event_accessors
' origin: languages/vb/tests/vb/test_vb_custom_event_accessors.rs

Class Publisher
    Private _count As Integer = 0
    
    ' Custom event with modifiers on accessors
    Public Custom Event Notify As EventHandler
        Private AddHandler(value As EventHandler)
            _count += 1
            Console.WriteLine("Added")
        End AddHandler
        RemoveHandler(value As EventHandler)
            _count -= 1
        End RemoveHandler
        RaiseEvent(sender As Object, e As EventArgs)
            Console.WriteLine("Raised " & _count)
        End RaiseEvent
    End Event
    
    Public Sub RegisterAndTrigger()
        AddHandler Notify, Sub(s, e) Console.WriteLine("Internal")
        RaiseEvent Notify(Me, EventArgs.Empty)
    End Sub
End Class

Module M
    Sub Main()
        Dim p As New Publisher()
        p.RegisterAndTrigger()
    End Sub
End Module
