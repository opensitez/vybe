' vybe-test: vb/vb_object_late_bound_method_call/test_vb_late_bound_inherited_method_call
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
    Class BaseClass
        Public Function BaseMethod() As String
            Return "BaseMethodResult"
        End Function
    End Class

    Class DerivedClass
        Inherits BaseClass
    End Class

    Sub Main()
        Dim obj As Object = New DerivedClass()
        __Check(CStr(CStr(obj.BaseMethod())), "BaseMethodResult")
    End Sub
End Module
