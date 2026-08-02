' vybe-test: vb/vb_spec_object_model/object_model_spec_overloaded_methods_dispatch_by_arity
' origin: languages/vb/tests/vb/test_vb_spec_object_model.rs

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

Class Printer
    Public Function Show() As String
        Return "zero"
    End Function
    Public Function Show(value As Integer) As String
        Return CStr(value)
    End Function
End Class
Module M
    Sub Main()
        Dim p As New Printer()
        __Check(CStr(p.Show()), "zero")
        __Check(CStr(p.Show(9)), "9")
    End Sub
End Module
