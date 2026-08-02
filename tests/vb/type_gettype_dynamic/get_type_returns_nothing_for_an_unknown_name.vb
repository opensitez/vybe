' vybe-test: vb/type_gettype_dynamic/get_type_returns_nothing_for_an_unknown_name
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

Module Program
    Sub Main()
        __Check(CStr(Type.GetType("NoSuchTypeAnywhere") Is Nothing), "True")
    End Sub
End Module
