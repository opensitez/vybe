' vybe-test: vb/vb_sorted_dictionary_keys_ordering/test_vb_sorted_dictionary_custom_key_comparer_descending
' origin: languages/vb/tests/vb/test_vb_sorted_dictionary_keys_ordering.rs

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

Imports System.Collections.Generic

Class DescendingComparer
    Implements IComparer(Of Integer)
    Public Function Compare(x As Integer, y As Integer) As Integer Implements IComparer(Of Integer).Compare
        Return y.CompareTo(x)
    End Function
End Class

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of Integer, String)(New DescendingComparer())
        dict(10) = "Ten"
        dict(30) = "Thirty"
        dict(20) = "Twenty"

        Dim keys = String.Join(",", dict.Keys)
        __Check(CStr(keys), "30,20,10")
    End Sub
End Module
