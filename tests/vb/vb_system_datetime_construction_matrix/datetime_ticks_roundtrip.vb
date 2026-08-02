' vybe-test: vb/vb_system_datetime_construction_matrix/datetime_ticks_roundtrip
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
        Dim source As DateTime = DateTime.Now
        Dim ticks As Long = source.Ticks
        Dim roundTrip As New DateTime(ticks)

        __Check(CStr(source.Ticks = roundTrip.Ticks), "True")
        __Check(CStr(source.Year = roundTrip.Year), "True")
    End Sub
End Module
