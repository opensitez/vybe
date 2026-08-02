' vybe-test: vb/vb_generic_interface_inheritance/test_vb_generic_interface_indexer_property
' origin: languages/vb/tests/vb/test_vb_generic_interface_inheritance.rs

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

Interface IGenericContainer(Of TKey, TValue)
    Default Property Item(key As TKey) As TValue
End Interface

Class SimpleMap
    Implements IGenericContainer(Of String, Integer)
    Default Public Property Item(key As String) As Integer Implements IGenericContainer(Of String, Integer).Item
        Get
            Return key.Length
        End Get
        Set(value As Integer)
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim container As IGenericContainer(Of String, Integer) = New SimpleMap()
        __Check(CStr(container("Hello")), "5")
    End Sub
End Module
