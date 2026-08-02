' vybe-test: vb/vb_date_serial_value/date_serial_value
' origin: languages/vb/tests/vb/test_vb_date_serial_value.rs

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
        ' DateSerial creates a date from year, month, day
        Dim d = DateSerial(2022, 10, 15)
        __Check(CStr(d.Year), "2022")
        __Check(CStr(d.Month), "10")
        
        ' DateValue parses a string
        Dim dv = DateValue("2023-05-20")
        __Check(CStr(dv.Day), "20")
    End Sub
End Module
