' vybe-test: vb/vb_sorted_dictionary_keys_ordering/test_vb_sorted_dictionary_custom_class_key_icomparable
' origin: languages/vb/tests/vb/test_vb_sorted_dictionary_keys_ordering.rs

Imports System
Imports System.Collections.Generic

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
        Console.WriteLine(String.Join(",", ids))
    End Sub
End Module
