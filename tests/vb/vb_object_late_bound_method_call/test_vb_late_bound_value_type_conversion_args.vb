' vybe-test: vb/vb_object_late_bound_method_call/test_vb_late_bound_value_type_conversion_args
' origin: languages/vb/tests/vb/test_vb_object_late_bound_method_call.rs

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
    Class ImplicitConverter
        Public Function DoubleVal(x As Double) As Double
            Return x * 2.0
        End Function
    End Class

    Sub Main()
        Dim obj As Object = New ImplicitConverter()
        ' Integer 10 should implicitly convert to Double 10.0 in late binding!
        Dim res As Double = CDbl(obj.DoubleVal(10))
        __Check(CStr(res), "20")
    End Sub
End Module
