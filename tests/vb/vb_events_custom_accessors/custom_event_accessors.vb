' vybe-test: vb/vb_events_custom_accessors/custom_event_accessors
' origin: languages/vb/tests/vb/test_vb_events_custom_accessors.rs

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

Class CustomEventSource
    ' Backing delegate
    Private ActionDelegate As Action
    
    ' Custom Event
    Public Custom Event ActionOccurred As Action
        AddHandler(value As Action)
            ActionDelegate = CType([Delegate].Combine(ActionDelegate, value), Action)
            __Check(CStr("Handler Added"), "Handler Added")
        End AddHandler
        
        RemoveHandler(value As Action)
            ActionDelegate = CType([Delegate].Remove(ActionDelegate, value), Action)
            __Check(CStr("Handler Removed"), "Raising Event")
        End RemoveHandler
        
        RaiseEvent()
            __Check(CStr("Raising Event"), "Action executed")
            If ActionDelegate IsNot Nothing Then
                ActionDelegate.Invoke()
            End If
        End RaiseEvent
    End Event
    
    Public Sub DoAction()
        RaiseEvent ActionOccurred()
    End Sub
End Class

Module M
    Sub OnAction()
        __Check(CStr("Action executed"), "Handler Removed")
    End Sub

    Sub Main()
        Dim source As New CustomEventSource()
        AddHandler source.ActionOccurred, AddressOf OnAction
        source.DoAction()
        RemoveHandler source.ActionOccurred, AddressOf OnAction
    End Sub
End Module
