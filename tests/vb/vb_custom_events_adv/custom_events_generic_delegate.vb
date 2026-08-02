' vybe-test: vb/vb_custom_events_adv/custom_events_generic_delegate
' origin: languages/vb/tests/vb/test_vb_custom_events_adv.rs

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

Class Subject
    Private _handlers As EventHandler(Of String)
    
    Public Custom Event Notify As EventHandler(Of String)
        AddHandler(value As EventHandler(Of String))
            _handlers = CType([Delegate].Combine(_handlers, value), EventHandler(Of String))
        End AddHandler
        RemoveHandler(value As EventHandler(Of String))
            _handlers = CType([Delegate].Remove(_handlers, value), EventHandler(Of String))
        End RemoveHandler
        RaiseEvent(sender As Object, e As String)
            If _handlers IsNot Nothing Then
                _handlers.Invoke(sender, e)
            End If
        End RaiseEvent
    End Event
    
    Public Sub Trigger(msg As String)
        RaiseEvent Notify(Me, msg)
    End Sub
End Class

Module M
    Sub OnNotify(sender As Object, e As String)
        __Check(CStr("Notified: " & e), "Notified: Hello")
    End Sub

    Sub Main()
        Dim s As New Subject()
        AddHandler s.Notify, AddressOf OnNotify
        s.Trigger("Hello")
    End Sub
End Module
