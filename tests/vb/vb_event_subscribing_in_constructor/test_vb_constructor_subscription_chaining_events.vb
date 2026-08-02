' vybe-test: vb/vb_event_subscribing_in_constructor/test_vb_constructor_subscription_chaining_events
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

Class ComponentA
    Public Event EventA As EventHandler
    Public Sub TriggerA()
        RaiseEvent EventA(Me, EventArgs.Empty)
    End Sub
End Class

Class ComponentB
    Public Event EventB As EventHandler
    Public Sub New(compA As ComponentA)
        AddHandler compA.EventA, Sub(s, e) RaiseEvent EventB(Me, EventArgs.Empty)
    End Sub
End Class

Module Program
    Sub Main()
        Dim ca As New ComponentA()
        Dim cb As New ComponentB(ca)
        AddHandler cb.EventB, Sub(s, e) __Check(CStr("Chained B Handled"), "Chained B Handled")
        ca.TriggerA()
    End Sub
End Module
