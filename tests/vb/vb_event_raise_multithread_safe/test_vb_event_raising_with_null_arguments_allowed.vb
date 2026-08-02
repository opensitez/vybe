' vybe-test: vb/vb_event_raise_multithread_safe/test_vb_event_raising_with_null_arguments_allowed
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

Class NullArgsEmitter
    Public Event CustomNotify As EventHandler
    Public Sub RaiseNullArgs()
        RaiseEvent CustomNotify(Nothing, Nothing)
    End Sub
End Class

Module Program
    Sub Main()
        Dim e As New NullArgsEmitter()
        AddHandler e.CustomNotify, Sub(sender, args)
            __Check(CStr("SenderIsNull=" & (sender Is Nothing) & "|ArgsIsNull=" & (args Is Nothing)), "SenderIsNull=True|ArgsIsNull=True")
        End Sub
        e.RaiseNullArgs()
    End Sub
End Module
