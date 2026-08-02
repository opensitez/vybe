' vybe-test: vb/vb_async_stream_iasyncenumerable/test_vb_async_iterator_yield_return
' origin: languages/vb/tests/vb/test_vb_async_stream_iasyncenumerable.rs

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

Imports System.Collections.Generic
Imports System.Threading.Tasks

Module Program
    Async Function GenerateNumbersAsync() As IAsyncEnumerable(Of Integer)
        ' Mock async stream pattern using list Task result
        Return FetchListAsync().Result
    End Function

    Async Function FetchListAsync() As Task(Of List(Of Integer))
        Await Task.Delay(10)
        Return New List(Of Integer) From {1, 2, 3}
    End Function

    Sub Main()
        Dim items = GenerateNumbersAsync().Result
        __Check(CStr(String.Join(",", items)), "1,2,3")
    End Sub
End Module
