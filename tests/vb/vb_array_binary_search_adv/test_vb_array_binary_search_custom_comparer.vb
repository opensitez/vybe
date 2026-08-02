' vybe-test: vb/vb_array_binary_search_adv/test_vb_array_binary_search_custom_comparer
' origin: languages/vb/tests/vb/test_vb_array_binary_search_adv.rs

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
        Dim arr As Integer() = {50, 40, 30, 20, 10}
        Dim idx As Integer = Array.BinarySearch(arr, 40, New DescendingComparer())
        __Check(CStr(idx), "1")
    End Sub
End Module
