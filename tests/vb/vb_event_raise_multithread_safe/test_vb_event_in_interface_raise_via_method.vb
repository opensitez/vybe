' vybe-test: vb/vb_event_raise_multithread_safe/test_vb_event_in_interface_raise_via_method
' origin: languages/vb/tests/vb/test_vb_event_raise_multithread_safe.rs

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

Interface IClickable
    Event Click As EventHandler
    Sub DoClick()
End Interface

Class LinkLabel
    Implements IClickable
    Public Event Click As EventHandler Implements IClickable.Click
    Public Sub DoClick() Implements IClickable.DoClick
        RaiseEvent Click(Me, EventArgs.Empty)
    End Sub
End Class

Module Program
    Sub Main()
        Dim c As IClickable = New LinkLabel()
        AddHandler c.Click, Sub(s, e) __Check(CStr("Link Clicked"), "Link Clicked")
        c.DoClick()
    End Sub
End Module
