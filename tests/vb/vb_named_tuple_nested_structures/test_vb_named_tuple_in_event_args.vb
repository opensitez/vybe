' vybe-test: vb/vb_named_tuple_nested_structures/test_vb_named_tuple_in_event_args
' origin: languages/vb/tests/vb/test_vb_named_tuple_nested_structures.rs

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

Class TaskRunner
    Public Event TaskProgress As Action(Of (StepName As String, Percent As Integer))
    Public Sub Report(name As String, pct As Integer)
        RaiseEvent TaskProgress((name, pct))
    End Sub
End Class

Module Program
    Sub Main()
        Dim runner As New TaskRunner()
        AddHandler runner.TaskProgress, Sub(info) __Check(CStr(info.StepName & " " & info.Percent & "%"), "Download 75%")
        runner.Report("Download", 75)
    End Sub
End Module
