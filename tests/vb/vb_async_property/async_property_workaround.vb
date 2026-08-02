' vybe-test: vb/vb_async_property/async_property_workaround
' origin: languages/vb/tests/vb/test_vb_async_property.rs

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

Class DataService
    ' Properties cannot be Async directly. 
    ' But they can return a Task(Of T).
    Public ReadOnly Property DataAsync As Task(Of String)
        Get
            Return FetchDataAsync()
        End Get
    End Property
    
    Private Async Function FetchDataAsync() As Task(Of String)
        Await Task.Delay(1)
        Return "Async Data"
    End Function
End Class

Module M
    Sub Main()
        Dim ds As New DataService()
        ' We synchronously wait for the task in Main
        Dim result = ds.DataAsync.Result
        __Check(CStr(result), "Async Data")
    End Sub
End Module
