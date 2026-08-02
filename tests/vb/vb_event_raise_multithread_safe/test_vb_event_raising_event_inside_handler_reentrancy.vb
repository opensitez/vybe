' vybe-test: vb/vb_event_raise_multithread_safe/test_vb_event_raising_event_inside_handler_reentrancy
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

Class ChainEmitter
    Public Event Step1 As Action
    Public Event Step2 As Action

    Public Sub Run()
        RaiseEvent Step1()
    End Sub
End Class

Module Program
    Sub Main()
        Dim c As New ChainEmitter()
        AddHandler c.Step1, Sub()
            __Check(CStr("Step1 Triggered"), "Step1 Triggered")
            RaiseEvent c.Step2()
        End Sub
        AddHandler c.Step2, Sub() __Check(CStr("Step2 Triggered"), "Step2 Triggered")
        c.Run()
    End Sub
End Module
