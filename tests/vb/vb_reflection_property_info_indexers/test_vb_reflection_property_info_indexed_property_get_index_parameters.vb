' vybe-test: vb/vb_reflection_property_info_indexers/test_vb_reflection_property_info_indexed_property_get_index_parameters
' origin: languages/vb/tests/vb/test_vb_reflection_property_info_indexers.rs

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

Imports System.Reflection

Class StringGrid
    Default Public Property Item(row As Integer, col As Integer) As String
        Get
            Return "R" & row & "C" & col
        End Get
        Set(value As String)
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim prop = GetType(StringGrid).GetProperty("Item")
        Dim indexParams = prop.GetIndexParameters()
        __Check(CStr(indexParams.Length & ":" & indexParams(0).Name & "," & indexParams(1).Name), "2:row,col")
    End Sub
End Module
