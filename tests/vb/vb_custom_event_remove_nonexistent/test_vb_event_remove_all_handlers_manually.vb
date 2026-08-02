' vybe-test: vb/vb_custom_event_remove_nonexistent/test_vb_event_remove_all_handlers_manually
' origin: languages/vb/tests/vb/test_vb_custom_event_remove_nonexistent.rs

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

Class MultiSubscriber
    Public Event Action As Action
    Public Sub Fire()
        RaiseEvent Action()
    End Sub
End Class

Module Program
    Private Sub A() : End Sub
    Private Sub B() : End Sub

    Sub Main()
        Dim ms As New MultiSubscriber()
        AddHandler ms.Action, AddressOf A
        AddHandler ms.Action, AddressOf B
        RemoveHandler ms.Action, AddressOf A
        RemoveHandler ms.Action, AddressOf B
        __Check(CStr("All removed safely"), "All removed safely")
        ms.Fire()
    End Sub
End Module
