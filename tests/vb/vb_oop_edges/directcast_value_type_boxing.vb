' vybe-test: vb/vb_oop_edges/directcast_value_type_boxing
' origin: languages/vb/tests/vb/test_vb_oop_edges.rs

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
        Dim x As Integer = 100
        Dim obj As Object = x
        ' DirectCast on boxed value type is exact type match only
        Dim y = DirectCast(obj, Integer)
        __Check(CStr(y), "100")
    End Sub
End Module
