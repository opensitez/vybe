' vybe-test: vb/vb_task_completion_source_set_result/test_vb_tcs_event_callback_resolution
' origin: languages/vb/tests/vb/test_vb_task_completion_source_set_result.rs

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
Imports System.Threading.Tasks

Class Producer
    Public Event DataAvailable As Action(Of String)
    Public Sub Produce(data As String)
        RaiseEvent DataAvailable(data)
    End Sub
End Class

Module Program
    Sub Main()
        Dim p As New Producer()
        Dim tcs As New TaskCompletionSource(Of String)()
        AddHandler p.DataAvailable, Sub(d) tcs.SetResult(d)

        p.Produce("Event Data")
        __Check(CStr(tcs.Task.Result), "Event Data")
    End Sub
End Module
