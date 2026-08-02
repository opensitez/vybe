' vybe-test: vb/vb_complex_class_hierarchy_generics/test_vb_nested_generic_class_instantiation
' origin: languages/vb/tests/vb/test_vb_complex_class_hierarchy_generics.rs

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

Class OuterTree(Of TKey)
    Class Node(Of TValue)
        Public Key As TKey
        Public Value As TValue
        Public Sub New(k As TKey, v As TValue)
            Key = k
            Value = v
        End Sub
    End Class
End Class

Module Program
    Sub Main()
        Dim node As New OuterTree(Of String).Node(Of Integer)("Key1", 100)
        __Check(CStr(node.Key & "=" & node.Value), "Key1=100")
    End Sub
End Module
