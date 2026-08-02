' vybe-test: vb/vb_byref_mutation/byref_datetime_day_can_advance_one_and_preserve_caller
' origin: languages/vb/tests/vb/test_vb_byref_mutation.rs

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
    Sub Advance(ByRef value As Date)
        value = value.AddDays(1)
    End Sub

    Sub Main()
        Dim value As Date = New Date(2024, 7, 1)
        Advance(value)
        __Check(CStr(value.Day.ToString()), "2")
    End Sub
End Module
