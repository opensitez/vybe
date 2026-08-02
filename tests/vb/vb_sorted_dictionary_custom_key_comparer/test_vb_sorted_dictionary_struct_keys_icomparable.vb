' vybe-test: vb/vb_sorted_dictionary_custom_key_comparer/test_vb_sorted_dictionary_struct_keys_icomparable
' origin: languages/vb/tests/vb/test_vb_sorted_dictionary_custom_key_comparer.rs

Imports System
Imports System.Collections.Generic

Structure Point
    Implements IComparable(Of Point)
    Public X As Integer
    Public Y As Integer
    Public Sub New(x As Integer, y As Integer)
        Me.X = x : Me.Y = y
    End Sub
    Public Function CompareTo(other As Point) As Integer Implements IComparable(Of Point).CompareTo
        Dim res = X.CompareTo(other.X)
        If res = 0 Then Return Y.CompareTo(other.Y)
        Return res
    End Function
End Structure

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of Point, String)()
        dict(New Point(2, 1)) = "P21"
        dict(New Point(1, 5)) = "P15"
        dict(New Point(1, 2)) = "P12"
        For Each kv In dict
            Console.WriteLine(kv.Key.X & "," & kv.Key.Y & "=" & kv.Value)
        Next
    End Sub
End Module
