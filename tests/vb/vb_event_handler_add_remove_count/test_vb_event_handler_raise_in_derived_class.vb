' vybe-test: vb/vb_event_handler_add_remove_count/test_vb_event_handler_raise_in_derived_class
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

Class BaseNotifier
    Public Event OnNotice As EventHandler
    Protected Sub RaiseNotice()
        RaiseEvent OnNotice(Me, EventArgs.Empty)
    End Sub
End Class

Class DerivedNotifier
    Inherits BaseNotifier
    Public Sub TriggerFromDerived()
        RaiseNotice()
    End Sub
End Class

Module Program
    Sub Main()
        Dim dn As New DerivedNotifier()
        AddHandler dn.OnNotice, Sub(s, e) __Check(CStr("Notice From Derived"), "Notice From Derived")
        dn.TriggerFromDerived()
    End Sub
End Module
