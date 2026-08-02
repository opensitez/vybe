' vybe-test: vb/vb_numeric_type_conversions_adv/test_vb_conv_tryparse_integer_valid
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
        Dim result As Integer
        Dim success As Boolean = Integer.TryParse("12345", result)
        __Check(CStr(success), "True")
        __Check(CStr(result), "12345")
    End Sub
End Module
