' vybe-test: vb/vb_numeric_type_conversions_adv/test_vb_conv_convert_change_type
' origin: languages/vb/tests/vb/test_vb_numeric_type_conversions_adv.rs

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
    Sub Main()
        Dim obj As Object = "99"
        Dim res As Object = Convert.ChangeType(obj, GetType(Integer))
        __Check(CStr(res.GetType().Name), "Int32")
        __Check(CStr(res), "99")
    End Sub
End Module
