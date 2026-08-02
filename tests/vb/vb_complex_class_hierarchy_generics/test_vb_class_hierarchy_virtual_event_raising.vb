' vybe-test: vb/vb_complex_class_hierarchy_generics/test_vb_class_hierarchy_virtual_event_raising
' origin: languages/vb/tests/vb/test_vb_complex_class_hierarchy_generics.rs

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

Class BaseEmitter
    Public Event Notice As EventHandler
    Protected Overridable Sub OnNotice()
        RaiseEvent Notice(Me, EventArgs.Empty)
    End Sub
    Public Sub Fire()
        OnNotice()
    End Sub
End Class

Class InterceptEmitter
    Inherits BaseEmitter
    Protected Overrides Sub OnNotice()
        __Check(CStr("Intercepted Before Fire"), "Intercepted Before Fire")
        MyBase.OnNotice()
    End Sub
End Class

Module Program
    Sub Main()
        Dim ie As New InterceptEmitter()
        AddHandler ie.Notice, Sub(s, e) __Check(CStr("Base Notice Fired"), "Base Notice Fired")
        ie.Fire()
    End Sub
End Module
