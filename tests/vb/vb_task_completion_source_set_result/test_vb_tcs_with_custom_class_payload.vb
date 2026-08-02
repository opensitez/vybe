' vybe-test: vb/vb_task_completion_source_set_result/test_vb_tcs_with_custom_class_payload
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

Imports System.Threading.Tasks

Class Response
    Public Status As String
    Public Sub New(s As String) : Status = s : End Sub
End Class

Module Program
    Sub Main()
        Dim tcs As New TaskCompletionSource(Of Response)()
        tcs.SetResult(New Response("200 OK"))
        __Check(CStr(tcs.Task.Result.Status), "200 OK")
    End Sub
End Module
