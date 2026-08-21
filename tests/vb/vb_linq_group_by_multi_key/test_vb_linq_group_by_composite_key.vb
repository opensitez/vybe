' vybe-test: vb/vb_linq_group_by_multi_key/test_vb_linq_group_by_composite_key
' origin: languages/vb/tests/vb/test_vb_linq_group_by_multi_key.rs

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

Imports System.Linq
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


Module Program
    Sub Main()
        Dim employees = {
        New With {.Dept = "IT", .Role = "Dev", .Name = "Alice"},
        New With {.Dept = "IT", .Role = "Dev", .Name = "Bob"},
        New With {.Dept = "IT", .Role = "QA", .Name = "Charlie"},
        New With {.Dept = "HR", .Role = "Recruiter", .Name = "David"}
        }

        Dim groups = From emp In employees
        Group emp By Key = New With {emp.Dept, emp.Role} Into Group

        __P(CStr(groups.Count()))
        For Each g In groups
            __P(CStr(g.Key.Dept & "-" & g.Key.Role & ":" & g.Group.Count()))
        Next
        __Check("4
IT-Dev:1
IT-Dev:1
IT-QA:1
HR-Recruiter:1")
    End Sub
End Module
