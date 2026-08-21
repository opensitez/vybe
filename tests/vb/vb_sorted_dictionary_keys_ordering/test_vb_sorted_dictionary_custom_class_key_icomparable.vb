' vybe-test: vb/vb_sorted_dictionary_keys_ordering/test_vb_sorted_dictionary_custom_class_key_icomparable
' origin: languages/vb/tests/vb/test_vb_sorted_dictionary_keys_ordering.rs

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
Imports System.Collections.Generic
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


Class EmployeeKey
    Implements IComparable(Of EmployeeKey)
    Public Id As Integer
    Public Sub New(i As Integer)
        Id = i
    End Sub
    Public Function CompareTo(other As EmployeeKey) As Integer Implements IComparable(Of EmployeeKey).CompareTo
        Return Id.CompareTo(other.Id)
    End Function
End Class

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of EmployeeKey, String)()
        dict(New EmployeeKey(50)) = "Fifty"
        dict(New EmployeeKey(10)) = "Ten"

        Dim ids As New List(Of Integer)()
        For Each k In dict.Keys
            ids.Add(k.Id)
        Next
        __P(CStr(String.Join(",", ids)))
        __Check("10,50")
    End Sub
End Module
