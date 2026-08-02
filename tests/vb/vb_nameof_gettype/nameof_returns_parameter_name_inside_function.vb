' vybe-test: vb/vb_nameof_gettype/nameof_returns_parameter_name_inside_function
' origin: languages/vb/tests/vb/test_vb_nameof_gettype.rs

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

Module M
    Function ShowName(value As Integer) As String
        Return NameOf(value)
    End Function

    Sub Main()
        __Check(CStr(ShowName(5)), "value")
    End Sub
End Module
