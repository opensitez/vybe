' vybe-test: vb/vb_generic_method_overloading/test_vb_generic_method_overload_by_arity
' origin: languages/vb/tests/vb/test_vb_generic_method_overloading.rs

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

Module Converter
    Public Function ConvertVal(Of T)(val As Object) As T
        Return CType(val, T)
    End Function

    Public Function ConvertVal(Of T1, T2)(val1 As Object, val2 As Object) As String
        Return val1.ToString() & "-" & val2.ToString()
    End Function
End Module

Module Program
    Sub Main()
        Dim i As Integer = Converter.ConvertVal(Of Integer)("100")
        Dim s As String = Converter.ConvertVal(Of Integer, String)(1, 2)
        __Check(CStr(i), "100")
        __Check(CStr(s), "1-2")
    End Sub
End Module
