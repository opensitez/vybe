' vybe-test: vb/vb_reflection_method_info_generic_invoke/test_vb_reflection_make_generic_method_two_type_args
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

Class Mapper
    Public Function Map(Of T1, T2)(item As T1, transform As System.Func(Of T1, T2)) As T2
        Return transform(item)
    End Function
End Class

Module Program
    Sub Main()
        Dim m As New Mapper()
        Dim openMethod = GetType(Mapper).GetMethod("Map")
        Dim closedMethod = openMethod.MakeGenericMethod(GetType(String), GetType(Integer))
        Dim fn As System.Func(Of String, Integer) = Function(s) s.Length
        Dim res = closedMethod.Invoke(m, {"VisualBasic", fn})
        __Check(CStr(res), "11")
    End Sub
End Module
