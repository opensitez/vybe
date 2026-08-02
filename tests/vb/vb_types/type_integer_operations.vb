' vybe-test: vb/vb_types/type_integer_operations
' origin: languages/vb/tests/vb/test_vb_types.rs

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
    Sub Main()
        Dim a As Integer = 10
        Dim b As Integer = 3
        __Check(CStr(a + b), "13")
        __Check(CStr(a - b), "7")
        __Check(CStr(a * b), "30")
        __Check(CStr(a \ b), "3")
        __Check(CStr(a Mod b), "1")
    End Sub
End Module
