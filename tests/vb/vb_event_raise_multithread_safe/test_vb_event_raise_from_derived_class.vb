' vybe-test: vb/vb_event_raise_multithread_safe/test_vb_event_raise_from_derived_class
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

Class BaseNotifier
    Public Event Notice As Action
    Protected Sub RaiseNotice()
        RaiseEvent Notice()
    End Sub
End Class

Class DerivedNotifier
    Inherits BaseNotifier
    Public Sub TriggerNotice()
        RaiseNotice()
    End Sub
End Class

Module Program
    Sub Main()
        Dim d As New DerivedNotifier()
        AddHandler d.Notice, Sub() __Check(CStr("Notice Triggered"), "Notice Triggered")
        d.TriggerNotice()
    End Sub
End Module
