' vybe-test: vb/vb_system_datetime_construction_matrix/datetime_from_parts_and_ticks
' origin: languages/vb/tests/vb/test_vb_system_datetime_construction_matrix.rs

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
        Dim dt As New DateTime(2026, 7, 21, 13, 40, 0)
        __Check(CStr(dt.Year), "2026")
        __Check(CStr(dt.Month), "7")
        __Check(CStr(dt.Day), "21")
        __Check(CStr(dt.Hour), "13")
        __Check(CStr(dt.Minute), "40")
    End Sub
End Module
