' vybe-test: vb/vb_system_datetime_construction_matrix/datetime_kind_and_utc_roundtrip
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
        Dim utc As DateTime = DateTime.SpecifyKind(New DateTime(2026, 7, 21, 0, 0, 0), DateTimeKind.Utc)
        Dim local As DateTime = utc.ToLocalTime()

        __Check(CStr(utc.Kind = DateTimeKind.Utc), "True")
        __Check(CStr(local.Kind <> DateTimeKind.Unspecified), "True")
    End Sub
End Module
