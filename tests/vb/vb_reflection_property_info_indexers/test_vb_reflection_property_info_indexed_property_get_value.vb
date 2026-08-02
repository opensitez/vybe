' vybe-test: vb/vb_reflection_property_info_indexers/test_vb_reflection_property_info_indexed_property_get_value
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

Class SimpleMap
    Default Public Property Item(key As String) As Integer
        Get
            Return key.Length
        End Get
        Set(value As Integer)
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim map As New SimpleMap()
        Dim prop = GetType(SimpleMap).GetProperty("Item")
        Dim val = prop.GetValue(map, {"VisualBasic"})
        __Check(CStr(val), "11")
    End Sub
End Module
