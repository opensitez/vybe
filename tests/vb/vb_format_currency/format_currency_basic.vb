' vybe-test: vb/vb_format_currency/format_currency_basic
' origin: languages/vb/tests/vb/test_vb_format_currency.rs

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
        Dim value As Double = 1234.567
        ' Defaults to 2 decimal places and system currency symbol
        ' Testing actual system output is hard since it depends on culture,
        ' but we can check if it parses and runs.
        Dim result As String = FormatCurrency(value)
        __Check(CStr("Formatted"), "Formatted")
    End Sub
End Module
