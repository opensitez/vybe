' vybe-test: vb/vb_conversion_methods/convert_type_code_returns_expected_type
' origin: languages/vb/tests/vb/test_vb_conversion_methods.rs

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

Imports System

Module M
    Sub Main()
        __Check(CStr(Convert.GetTypeCode("text") = TypeCode.String), "True")
        __Check(CStr(Convert.GetTypeCode(1) = TypeCode.Int32), "True")
    End Sub
End Module
