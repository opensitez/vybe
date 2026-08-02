' vybe-test: vb/vb_info_isnumeric_isdate/info_isnumeric_isdate
' origin: languages/vb/tests/vb/test_vb_info_isnumeric_isdate.rs

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
        ' IsNumeric
        __Check(CStr(IsNumeric("123")), "True")
        __Check(CStr(IsNumeric("12.34")), "True")
        __Check(CStr(IsNumeric("abc")), "False")
        
        ' IsDate
        __Check(CStr(IsDate("2023-01-01")), "True")
        __Check(CStr(IsDate("Not A Date")), "False")
    End Sub
End Module
