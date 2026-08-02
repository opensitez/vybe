' vybe-test: vb/vb_async_interfaces/async_interfaces
' origin: languages/vb/tests/vb/test_vb_async_interfaces.rs

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

Interface IWorker
    Function DoWorkAsync() As Task(Of Integer)
End Interface

Class Worker
    Implements IWorker
    
    ' The Async modifier goes on the implementation, not the interface
    Public Async Function DoWorkAsync() As Task(Of Integer) Implements IWorker.DoWorkAsync
        Await Task.Delay(1)
        Return 42
    End Function
End Class

Module M
    Sub Main()
        Dim w As IWorker = New Worker()
        Dim t As Task(Of Integer) = w.DoWorkAsync()
        t.Wait()
        __Check(CStr(t.Result), "42")
    End Sub
End Module
