' vybe-test: vb/vb_object_late_bound_method_call/test_vb_late_bound_method_raising_exception
' origin: languages/vb/tests/vb/test_vb_object_late_bound_method_call.rs

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

Module Program
    Class Faulty
        Public Sub Fail()
            Throw New InvalidOperationException("Custom Fault")
        End Sub
    End Class

    Sub Main()
        Dim obj As Object = New Faulty()
        Try
            obj.Fail()
        Catch ex As InvalidOperationException
            __Check(CStr(ex.Message), "Custom Fault")
        End Try
    End Sub
End Module
