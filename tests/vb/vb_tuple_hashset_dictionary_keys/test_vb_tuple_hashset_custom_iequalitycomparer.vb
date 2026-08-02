' vybe-test: vb/vb_tuple_hashset_dictionary_keys/test_vb_tuple_hashset_custom_iequalitycomparer
' origin: languages/vb/tests/vb/test_vb_tuple_hashset_dictionary_keys.rs

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
Imports System.Collections.Generic

Class TupleIgnoreCaseComparer
    Implements IEqualityComparer(Of (String, Integer))
    Public Function Equals(x As (String, Integer), y As (String, Integer)) As Boolean Implements IEqualityComparer(Of (String, Integer)).Equals
        Return String.Equals(x.Item1, y.Item1, StringComparison.OrdinalIgnoreCase) AndAlso x.Item2 = y.Item2
    End Function
    Public Function GetHashCode(obj As (String, Integer)) As Integer Implements IEqualityComparer(Of (String, Integer)).GetHashCode
        Return StringComparer.OrdinalIgnoreCase.GetHashCode(obj.Item1) Xor obj.Item2.GetHashCode()
    End Function
End Class

Module Program
    Sub Main()
        Dim set As New HashSet(Of (String, Integer))(New TupleIgnoreCaseComparer())
        set.Add(("apple", 10))
        set.Add(("APPLE", 10))
        __Check(CStr(set.Count), "1")
    End Sub
End Module
