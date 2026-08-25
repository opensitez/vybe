' vybe-test: vb/vb_system_datetime_offset_matrix/datetime_offset_compare_to_operator
' origin: languages/vb/tests/vb/test_vb_system_datetime_offset_matrix.rs

' Vybe test harness — Visual Basic.
'
' Real VB source alongside harness/go/check.go and harness/js/check.js, the way
' test262's assert.js is JavaScript.
'
' A test's verdict is its EXIT CODE. __Check prints its diagnostic BEFORE
' throwing: an uncaught exception surfaces as `RuntimeError: [object]`, which
' says nothing at all.
'
' Output is COLLECTED, not paired. The emitter rewrites every
' `Console.WriteLine(x)` into `__P(CStr(x))` and compares the whole output once
' at the end of `Sub Main`. Pairing the i-th print with the i-th expected line
' cannot assert anything about a loop, and loops alone were 402 of VB's 6,671
' cases.
'
' Rendering happens at the CALL SITE via `CStr`, where the expression still has
' its static type — the same reason the C# harness renders with `.ToString()`
' rather than inside the helper.

Imports System
Module VybeCheck
    Public __buf As String = ""

    Sub __P(s As String)
        __buf = __buf & s & vbLf
    End Sub

    Sub __Pr(s As String)
        __buf = __buf & s
    End Sub

    ' The final WriteLine contributes a trailing newline that the expected line
    ' vector never carried, so BOTH forms are accepted.
    Sub __Check(want As String)
        If __buf <> want AndAlso __buf <> want & vbLf Then
            Console.WriteLine("FAIL: want [" & want & "] got [" & __buf & "]")
            Throw New Exception("assertion failed")
        End If
    End Sub
End Module


Module M
    Sub Main()
        ' The two must denote the SAME INSTANT for the assertions below.
        ' As written (`0,0,0` at +01:00 vs `1,0,0` at +00:00) they were
        ' 2023-12-31T23:00Z and 2024-01-01T01:00Z — two hours apart — and real
        ' VB.NET answers `False`/`-1`, not `True`/`0` (checked with
        ' `tools/vbrun`). `a` moved to 02:00+01:00, which IS 01:00Z.
        Dim a As New DateTimeOffset(2024, 1, 1, 2, 0, 0, TimeSpan.FromHours(1))
        Dim b As New DateTimeOffset(2024, 1, 1, 1, 0, 0, TimeSpan.Zero)
        __P(CStr(a = b))
        __P(CStr(a.CompareTo(b)))
        __Check("True
0")
    End Sub
End Module
