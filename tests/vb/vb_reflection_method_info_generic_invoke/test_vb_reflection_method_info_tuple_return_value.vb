' vybe-test: vb/vb_reflection_method_info_generic_invoke/test_vb_reflection_method_info_tuple_return_value
' origin: languages/vb/tests/vb/test_vb_reflection_method_info_generic_invoke.rs

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

Class TupleService
    Public Function GetPair() As (Code As Integer, Name As String)
        Return (200, "OK")
    End Function
End Class

Module Program
    Sub Main()
        Dim svc As New TupleService()
        Dim m = GetType(TupleService).GetMethod("GetPair")
        Dim res As (Integer, String) = CType(m.Invoke(svc, Nothing), (Integer, String))
        __Check(CStr(res.Item1 & " " & res.Item2), "200 OK")
    End Sub
End Module
