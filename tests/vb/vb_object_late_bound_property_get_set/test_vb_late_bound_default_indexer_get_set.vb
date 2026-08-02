' vybe-test: vb/vb_object_late_bound_property_get_set/test_vb_late_bound_default_indexer_get_set
' origin: languages/vb/tests/vb/test_vb_object_late_bound_property_get_set.rs

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

Module Program
    Class CustomDictionary
        Private storage As New System.Collections.Generic.Dictionary(Of String, String)()
        Default Public Property Item(key As String) As String
            Get
                Return storage(key)
            End Get
            Set(value As String)
                storage(key) = value
            End Set
        End Property
    End Class

    Sub Main()
        Dim obj As Object = New CustomDictionary()
        obj("host") = "localhost"
        Dim host As String = CStr(obj("host"))
        __Check(CStr(host), "localhost")
    End Sub
End Module
