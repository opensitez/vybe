' vybe-test: vb/vb_system_array_2d_matrix/array_2d_matrix_sum_rows_and_columns
' origin: languages/vb/tests/vb/test_vb_system_array_2d_matrix.rs

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
        Dim m(2, 1) As Integer
        m(0, 0) = 1
        m(0, 1) = 2
        m(1, 0) = 3
        m(1, 1) = 4
        m(2, 0) = 5
        m(2, 1) = 6

        Dim rowSums(2) As Integer
        For r As Integer = m.GetLowerBound(0) To m.GetUpperBound(0)
            For c As Integer = m.GetLowerBound(1) To m.GetUpperBound(1)
                rowSums(r) += m(r, c)
            Next
        Next

        __P(CStr(rowSums(0)))
        __P(CStr(rowSums(1)))
        __P(CStr(rowSums(2)))

        Dim col0 As Integer = m(0, 0) + m(1, 0) + m(2, 0)
        Dim col1 As Integer = m(0, 1) + m(1, 1) + m(2, 1)
        __P(CStr(col0))
        __P(CStr(col1))
        __Check("3
7
11
9
12")
    End Sub
End Module
