' vybe-test: vb/type_gettype_dynamic/get_type_resolves_a_user_declared_class
' origin: languages/vb/tests/vb/test_type_gettype_dynamic.rs

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

Class Widget
End Class

Module Program
    Sub Main()
        __Check(CStr(Type.GetType("Widget") IsNot Nothing), "True")
    End Sub
End Module
