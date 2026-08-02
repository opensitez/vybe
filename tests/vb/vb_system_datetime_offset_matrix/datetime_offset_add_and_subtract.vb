' vybe-test: vb/vb_system_datetime_offset_matrix/datetime_offset_add_and_subtract
' origin: languages/vb/tests/vb/test_vb_system_datetime_offset_matrix.rs

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

Imports System

Module M
    Sub Main()
        Dim base As New DateTimeOffset(2024, 1, 1, 0, 0, 0, TimeSpan.Zero)
        Dim later As DateTimeOffset = base.AddDays(1).AddHours(3)
        Dim earlier As DateTimeOffset = later.AddHours(-3)
        __Check(CStr(later > base), "True")
        __Check(CStr(earlier = base), "True")
    End Sub
End Module
