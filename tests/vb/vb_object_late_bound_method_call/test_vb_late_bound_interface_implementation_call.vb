' vybe-test: vb/vb_object_late_bound_method_call/test_vb_late_bound_interface_implementation_call
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

Module Program
    Interface IService
        Function Execute() As String
    End Interface

    Class ServiceImpl
        Implements IService
        Public Function Execute() As String Implements IService.Execute
            Return "Executed"
        End Function
    End Class

    Sub Main()
        Dim obj As Object = New ServiceImpl()
        __Check(CStr(CStr(obj.Execute())), "Executed")
    End Sub
End Module
