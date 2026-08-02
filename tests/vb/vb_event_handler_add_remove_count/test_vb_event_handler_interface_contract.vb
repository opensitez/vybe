' vybe-test: vb/vb_event_handler_add_remove_count/test_vb_event_handler_interface_contract
' origin: languages/vb/tests/vb/test_vb_event_handler_add_remove_count.rs

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
End Interface

Class ButtonWidget
    Implements IClickable
    Public Event Click As EventHandler Implements IClickable.Click

    Public Sub ClickMe()
        RaiseEvent Click(Me, EventArgs.Empty)
    End Sub
End Class

Module Program
    Sub Main()
        Dim widget As IClickable = New ButtonWidget()
        AddHandler widget.Click, Sub(s, e) __Check(CStr("Interface Click Handled"), "Interface Click Handled")
        CType(widget, ButtonWidget).ClickMe()
    End Sub
End Module
