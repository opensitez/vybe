' vybe-test: vb/vb_operators_gettype/operator_gettype
' origin: languages/vb/tests/vb/test_vb_operators_gettype.rs

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
        ' GetType returns a System.Type object for the specified type
        Dim t As Type = GetType(String)
        __Check(CStr(t.Name), "String")
        
        Dim t2 As Type = GetType(Integer)
        __Check(CStr(t2.Name), "Int32")
    End Sub
End Module
