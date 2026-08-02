' vybe-test: vb/vb_interface_hierarchy_multiple_inheritance/test_vb_interface_indexer_multiple_overloads
' origin: languages/vb/tests/vb/test_vb_interface_hierarchy_multiple_inheritance.rs

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

Interface IIndexable
    Default Property Item(key As String) As String
    Default Property Item(index As Integer) As String
End Interface

Class DictionaryAdapter
    Implements IIndexable
    Public Property Item(key As String) As String Implements IIndexable.Item
        Get
            Return "Key_" & key
        End Get
        Set(value As String)
        End Set
    End Property
    Public Property Item(index As Integer) As String Implements IIndexable.Item
        Get
            Return "Idx_" & index
        End Get
        Set(value As String)
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim idx As IIndexable = New DictionaryAdapter()
        __Check(CStr(idx("name") & "|" & idx(42)), "Key_name|Idx_42")
    End Sub
End Module
