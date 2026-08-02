' vybe-test: vb/vb_custom_event_blocks/custom_event_blocks
' origin: languages/vb/tests/vb/test_vb_custom_event_blocks.rs

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

Class EventPublisher
    ' A Custom Event allows defining AddHandler, RemoveHandler, and RaiseEvent blocks
    Public Custom Event MyEvent As EventHandler
        AddHandler(value As EventHandler)
            __Check(CStr("Handler Added"), "Handler Added")
        End AddHandler
        RemoveHandler(value As EventHandler)
            __Check(CStr("Handler Removed"), "Event Raised")
        End RemoveHandler
        RaiseEvent(sender As Object, e As EventArgs)
            __Check(CStr("Event Raised"), "Handler Removed")
        End RaiseEvent
    End Event
    
    Public Sub Trigger()
        RaiseEvent MyEvent(Me, EventArgs.Empty)
    End Sub
End Class

Module M
    Sub Handler(sender As Object, e As EventArgs)
    End Sub

    Sub Main()
        Dim p As New EventPublisher()
        AddHandler p.MyEvent, AddressOf Handler
        p.Trigger()
        RemoveHandler p.MyEvent, AddressOf Handler
    End Sub
End Module
