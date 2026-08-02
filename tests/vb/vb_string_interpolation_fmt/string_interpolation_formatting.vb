' vybe-test: vb/vb_string_interpolation_fmt/string_interpolation_formatting
' origin: languages/vb/tests/vb/test_vb_string_interpolation_fmt.rs

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

Imports System.Globalization

Module M
    Sub Main()
        Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture
        
        Dim price As Decimal = 12.345D
        Dim pct As Double = 0.75
        
        ' Interpolation with formatting
        __Check(CStr($"Price: {price:F2}"), "Price: 12.35")
        __Check(CStr($"Percent: {pct:P0}"), "Percent: 75 %")
        
        ' Interpolation with alignment
        __Check(CStr($"[{price,10:F1}]"), "[      12.3]")
        __Check(CStr($"[{price,-10:F1}]"), "[12.3      ]")
    End Sub
End Module
