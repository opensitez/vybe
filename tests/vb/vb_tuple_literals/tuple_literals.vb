' vybe-test: vb/vb_tuple_literals/tuple_literals
' origin: languages/vb/tests/vb/test_vb_tuple_literals.rs

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
        ' Value Tuple literals in VB 15
        Dim t1 = (1, "A")
        __Check(CStr(t1.Item1), "1")
        __Check(CStr(t1.Item2), "A")
        
        ' Named tuple elements
        Dim t2 = (Id:=2, Name:="B")
        __Check(CStr(t2.Id), "2")
        __Check(CStr(t2.Name), "B")
    End Sub
End Module
