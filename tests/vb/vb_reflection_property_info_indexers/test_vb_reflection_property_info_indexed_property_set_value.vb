' vybe-test: vb/vb_reflection_property_info_indexers/test_vb_reflection_property_info_indexed_property_set_value
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

Imports System.Collections.Generic

Class Cache
    Private store As New Dictionary(Of String, String)()
    Default Public Property Item(key As String) As String
        Get
            Return store(key)
        End Get
        Set(value As String)
            store(key) = value
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim c As New Cache()
        Dim prop = GetType(Cache).GetProperty("Item")
        prop.SetValue(c, "Data100", {"K1"})
        __Check(CStr(c("K1")), "Data100")
    End Sub
End Module
