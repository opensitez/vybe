' vybe-test: vb/vb_sorted_dictionary_custom_key_comparer/test_vb_sorted_dictionary_custom_length_comparer
' origin: languages/vb/tests/vb/test_vb_sorted_dictionary_custom_key_comparer.rs

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

Class StringLengthComparer
    Implements IComparer(Of String)
    Public Function Compare(x As String, y As String) As Integer Implements IComparer(Of String).Compare
        Dim res = x.Length.CompareTo(y.Length)
        If res = 0 Then Return x.CompareTo(y)
        Return res
    End Function
End Class

Module Program
    Sub Main()
        Dim dict As New SortedDictionary(Of String, Integer)(New StringLengthComparer())
        dict("elephant") = 8
        dict("cat") = 3
        dict("dog") = 3
        __Check(CStr(String.Join(",", dict.Keys)), "cat,dog,elephant")
    End Sub
End Module
