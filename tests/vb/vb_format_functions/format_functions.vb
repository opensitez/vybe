' vybe-test: vb/vb_format_functions/format_functions
' origin: languages/vb/tests/vb/test_vb_format_functions.rs

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
        ' Format functions
        __Check(CStr(Format(12.34, "0.0")), "12.3")
        
        ' In some locales currency symbol changes, so we just check it doesn't throw
        Dim currencyStr = FormatCurrency(12.34)
        __Check(CStr(currencyStr.Length > 0), "True")
        
        Dim numStr = FormatNumber(12.34, 1)
        __Check(CStr(numStr.Length > 0), "True")
        
        Dim pctStr = FormatPercent(0.123, 1)
        __Check(CStr(pctStr.Length > 0), "True")
    End Sub
End Module
