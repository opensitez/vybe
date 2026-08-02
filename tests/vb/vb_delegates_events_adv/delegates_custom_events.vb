' vybe-test: vb/vb_delegates_events_adv/delegates_custom_events
' origin: languages/vb/tests/vb/test_vb_delegates_events_adv.rs

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

Class Timer
    Private _tickHandlers As System.EventHandler
    
    ' Custom event accessors
    Public Custom Event Tick As System.EventHandler
        AddHandler(value As System.EventHandler)
            _tickHandlers = CType([Delegate].Combine(_tickHandlers, value), System.EventHandler)
            __Check(CStr("Handler added"), "Handler added")
        End AddHandler
        RemoveHandler(value As System.EventHandler)
            _tickHandlers = CType([Delegate].Remove(_tickHandlers, value), System.EventHandler)
            __Check(CStr("Handler removed"), "Tick occurred")
        End RemoveHandler
        RaiseEvent(sender As Object, e As System.EventArgs)
            If _tickHandlers IsNot Nothing Then
                _tickHandlers.Invoke(sender, e)
            End If
        End RaiseEvent
    End Event
    
    Public Sub DoTick()
        RaiseEvent Tick(Me, System.EventArgs.Empty)
    End Sub
End Class

Module M
    Sub OnTick(sender As Object, e As System.EventArgs)
        __Check(CStr("Tick occurred"), "Handler removed")
    End Sub

    Sub Main()
        Dim t As New Timer()
        AddHandler t.Tick, AddressOf OnTick
        t.DoTick()
        RemoveHandler t.Tick, AddressOf OnTick
    End Sub
End Module
