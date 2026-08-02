' vybe-test: vb/vb_system_datetime_construction_matrix/datetime_adds_and_comparisons_are_expected
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
        Dim base As DateTime = New DateTime(2026, 1, 1)
        Dim plusDays As DateTime = base.AddDays(1)
        Dim plusMonths As DateTime = base.AddMonths(1)

        __Check(CStr(base < plusDays), "True")
        __Check(CStr(plusDays.Day), "2")
        __Check(CStr(plusMonths.Month), "2")
    End Sub
End Module
