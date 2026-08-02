' vybe-test: vb/vb_indexer_multidimensional/test_vb_indexer_overloaded_string_and_int
' origin: languages/vb/tests/vb/test_vb_indexer_multidimensional.rs

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

Class DataStore
    Private _byInt As New System.Collections.Generic.Dictionary(Of Integer, String)()
    Private _byStr As New System.Collections.Generic.Dictionary(Of String, String)()

    Default Public Property Item(id As Integer) As String
        Get
            Return _byInt(id)
        End Get
        Set(value As String)
            _byInt(id) = value
        End Set
    End Property

    Default Public Property Item(key As String) As String
        Get
            Return _byStr(key)
        End Get
        Set(value As String)
            _byStr(key) = value
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim ds As New DataStore()
        ds(1) = "NumOne"
        ds("A") = "StrA"
        __Check(CStr(ds(1)), "NumOne")
        __Check(CStr(ds("A")), "StrA")
    End Sub
End Module
