' vybe-test: vb/vb_index_out_of_range_exception/test_vb_safe_array_element_getter_extension
' origin: languages/vb/tests/vb/test_vb_index_out_of_range_exception.rs

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

Imports System.Runtime.CompilerServices

Module ArrayExtensions
    <Extension()>
    Public Function ElementAtOrDefault(Of T)(arr As T(), index As Integer, defaultValue As T) As T
        If arr IsNot Nothing AndAlso index >= 0 AndAlso index < arr.Length Then
            Return arr(index)
        End If
        Return defaultValue
    End Function
End Module

Module Program
    Sub Main()
        Dim arr As Integer() = {10, 20, 30}
        __Check(CStr(arr.ElementAtOrDefault(1, -1) & "|" & arr.ElementAtOrDefault(5, -1)), "20|-1")
    End Sub
End Module
