' vybe-test: vb/vb_linq_join_inner_outer/test_vb_linq_inner_join_query_syntax
' origin: languages/vb/tests/vb/test_vb_linq_join_inner_outer.rs

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
        Dim departments = {
            New With {.Id = 1, .Name = "Engineering"},
            New With {.Id = 2, .Name = "Marketing"}
        }

        Dim employees = {
            New With {.Name = "Alice", .DeptId = 1},
            New With {.Name = "Bob", .DeptId = 1},
            New With {.Name = "Charlie", .DeptId = 2}
        }

        Dim query = From emp In employees
                    Join dept In departments On emp.DeptId Equals dept.Id
                    Select emp.Name, dept.Name

        For Each item In query
            __P(CStr(item.Name & " in " & item.dept_Name))
        Next
        __Check("Alice in Engineering
Bob in Engineering
Charlie in Marketing")
    End Sub
End Module
