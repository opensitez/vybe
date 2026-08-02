' vybe-test: vb/vb_event_raise_multithread_safe/test_vb_event_custom_args_mutation_in_handler
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

Class CancelEventArgs
    Inherits EventArgs
    Public Property Cancel As Boolean = False
End Class

Class Worker
    Public Event QueryCancel As EventHandler(Of CancelEventArgs)
    Public Function TryPerformWork() As Boolean
        Dim args As New CancelEventArgs()
        RaiseEvent QueryCancel(Me, args)
        Return Not args.Cancel
    End Function
End Class

Module Program
    Sub Main()
        Dim w As New Worker()
        AddHandler w.QueryCancel, Sub(sender, e) e.Cancel = True
        __Check(CStr("Can Work: " & w.TryPerformWork()), "Can Work: False")
    End Sub
End Module
