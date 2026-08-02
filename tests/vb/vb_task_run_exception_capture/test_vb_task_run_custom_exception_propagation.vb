' vybe-test: vb/vb_task_run_exception_capture/test_vb_task_run_custom_exception_propagation
' origin: languages/vb/tests/vb/test_vb_task_run_exception_capture.rs

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

Class CustomDomainException
    Inherits Exception
    Public ErrorCode As Integer
    Public Sub New(code As Integer, msg As String)
        MyBase.New(msg)
        ErrorCode = code
    End Sub
End Class

Module Program
    Sub Main()
        Dim t = Task.Run(Sub()
            Throw New CustomDomainException(404, "Entity Not Found")
        End Sub)
        Try
            t.Wait()
        Catch ex As AggregateException
            Dim cust = CType(ex.InnerException, CustomDomainException)
            __Check(CStr(cust.ErrorCode & ": " & cust.Message), "404: Entity Not Found")
        End Try
    End Sub
End Module
