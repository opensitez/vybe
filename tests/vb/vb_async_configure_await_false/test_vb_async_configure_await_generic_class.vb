' vybe-test: vb/vb_async_configure_await_false/test_vb_async_configure_await_generic_class
' origin: languages/vb/tests/vb/test_vb_async_configure_await_false.rs

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

Class AsyncWorker(Of T)
    Public Async Function ProcessAsync(input As T) As Task(Of T)
        Await Task.Delay(5).ConfigureAwait(False)
        Return input
    End Function
End Class

Module Program
    Sub Main()
        Dim w As New AsyncWorker(Of String)()
        Dim t = w.ProcessAsync("GenericInput")
        __Check(CStr(t.Result), "GenericInput")
    End Sub
End Module
