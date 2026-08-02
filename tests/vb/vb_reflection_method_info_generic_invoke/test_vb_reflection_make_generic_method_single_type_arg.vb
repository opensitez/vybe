' vybe-test: vb/vb_reflection_method_info_generic_invoke/test_vb_reflection_make_generic_method_single_type_arg
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

Class GenericCalculator
    Public Function Identity(Of T)(val As T) As T
        Return val
    End Function
End Class

Module Program
    Sub Main()
        Dim calc As New GenericCalculator()
        Dim openMethod = GetType(GenericCalculator).GetMethod("Identity")
        Dim closedMethod = openMethod.MakeGenericMethod(GetType(Integer))
        Dim res = closedMethod.Invoke(calc, {42})
        __Check(CStr(res.GetType().Name & "=" & res), "Int32=42")
    End Sub
End Module
