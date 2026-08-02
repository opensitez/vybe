' vybe-test: vb/vb_sorted_dictionary_operations/test_vb_sorted_dict_custom_comparer
' origin: languages/vb/tests/vb/test_vb_sorted_dictionary_operations.rs

Imports System.Collections.Generic

Class DescendingIntComparer
    Implements IComparer(Of Integer)
    Public Function Compare(x As Integer, y As Integer) As Integer Implements IComparer(Of Integer).Compare
        Return y.CompareTo(x)
    End Function
End Class

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of Integer, String)(New DescendingIntComparer())
        dict(1) = "One"
        dict(3) = "Three"
        dict(2) = "Two"
        For Each kvp In dict
            Console.WriteLine(kvp.Key)
        Next
    End Sub
End Module
