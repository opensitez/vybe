' vybe-test: vb/vb_event_subscribing_in_constructor/test_vb_with_events_constructor_wireup
' origin: languages/vb/tests/vb/test_vb_event_subscribing_in_constructor.rs

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

Class Clock
    Public Event Tick As EventHandler
    Public Sub Start()
        RaiseEvent Tick(Me, EventArgs.Empty)
    End Sub
End Class

Class ClockListener
    Private WithEvents myClock As Clock

    Public Sub New(c As Clock)
        myClock = c ' Assigning WithEvents field automatically wires Handles methods!
    End Sub

    Private Sub OnTick(sender As Object, e As EventArgs) Handles myClock.Tick
        __Check(CStr("Clock Ticked via WithEvents"), "Clock Ticked via WithEvents")
    End Sub
End Class

Module Program
    Sub Main()
        Dim c As New Clock()
        Dim listener As New ClockListener(c)
        c.Start()
    End Sub
End Module
