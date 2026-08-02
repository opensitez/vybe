' vybe-test: vb/vb_interface_hierarchy_multiple_inheritance/test_vb_interface_event_conflict_disambiguation
' origin: languages/vb/tests/vb/test_vb_interface_hierarchy_multiple_inheritance.rs

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

Interface IEventA
    Event OnEvent As EventHandler
End Interface

Interface IEventB
    Event OnEvent As EventHandler
End Interface

Class Dispatcher
    Implements IEventA, IEventB
    Public Event EventA As EventHandler Implements IEventA.OnEvent
    Public Event EventB As EventHandler Implements IEventB.OnEvent

    Public Sub RaiseA()
        RaiseEvent EventA(Me, EventArgs.Empty)
    End Sub
    Public Sub RaiseB()
        RaiseEvent EventB(Me, EventArgs.Empty)
    End Sub
End Class

Module Program
    Sub Main()
        Dim d As New Dispatcher()
        Dim ea As IEventA = d
        Dim eb As IEventB = d

        AddHandler ea.OnEvent, Sub(s, e) __Check(CStr("A Raised"), "A Raised")
        AddHandler eb.OnEvent, Sub(s, e) __Check(CStr("B Raised"), "B Raised")

        d.RaiseA()
        d.RaiseB()
    End Sub
End Module
