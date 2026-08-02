' vybe-test: vb/vb_system_linq_ordering_matrix/linq_ordering_secondary_by_key
' origin: languages/vb/tests/vb/test_vb_system_linq_ordering_matrix.rs

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

Class Item
    Public Name As String
    Public Score As Integer

    Public Sub New(name As String, score As Integer)
        Me.Name = name
        Me.Score = score
    End Sub
End Class

Module M
    Sub Main()
        Dim data = {
            New Item("c", 1),
            New Item("a", 2),
            New Item("b", 2),
            New Item("a", 1)
        }

        Dim sorted = data.OrderBy(Function(i) i.Score).ThenBy(Function(i) i.Name)
        Dim firstName As String = sorted(0).Name
        Dim lastName As String = sorted.Last().Name

        __Check(CStr(sorted.First().Score), "1")
        __Check(CStr(firstName), "a")
        __Check(CStr(lastName), "c")
    End Sub
End Module
