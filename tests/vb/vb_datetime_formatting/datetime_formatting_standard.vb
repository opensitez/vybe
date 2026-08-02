' vybe-test: vb/vb_datetime_formatting/datetime_formatting_standard
' origin: languages/vb/tests/vb/test_vb_datetime_formatting.rs

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
        Dim d As Date = New Date(2024, 1, 1, 15, 30, 0)
        
        ' Ensure invariant culture for consistent results across environments
        Thread.CurrentThread.CurrentCulture = CultureInfo.InvariantCulture
        
        __Check(CStr(d.ToString("yyyy-MM-dd")), "2024-01-01")
        __Check(CStr(d.ToString("HH:mm:ss")), "15:30:00")
        __Check(CStr(d.ToString("yyyy-MM-dd HH:mm:ss")), "2024-01-01 15:30:00")
    End Sub
End Module
