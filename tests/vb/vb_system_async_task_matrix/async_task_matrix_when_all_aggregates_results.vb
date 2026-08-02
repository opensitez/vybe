' vybe-test: vb/vb_system_async_task_matrix/async_task_matrix_when_all_aggregates_results
' origin: languages/vb/tests/vb/test_vb_system_async_task_matrix.rs

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
Imports System.Linq
Imports System.Threading.Tasks

Module M
    Sub Main()
        Dim a As Task(Of Integer) = Task.Run(Function() 1)
        Dim b As Task(Of Integer) = Task.Run(Function() 2)
        Dim c As Task(Of Integer) = Task.Run(Function() 3)

        Dim all As Task(Of Integer()) = Task.WhenAll(a, b, c)
        __Check(CStr(all.IsCompleted), "True")
        __Check(CStr(all.Result.Sum()), "6")
        __Check(CStr(all.Result.Length), "3")
    End Sub
End Module
