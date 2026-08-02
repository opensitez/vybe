' vybe-test: vb/vb_property_indexed_multi_parameter/test_vb_indexed_property_derived_class_override
' origin: languages/vb/tests/vb/test_vb_property_indexed_multi_parameter.rs

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

Class BaseStore
    Default Public Overridable Property Item(id As Integer) As String
        Get
            Return "BaseStore"
        End Get
        Set(value As String)
        End Set
    End Property
End Class

Class CustomStore
    Inherits BaseStore
    Default Public Overrides Property Item(id As Integer) As String
        Get
            Return "CustomStore_" & id
        End Get
        Set(value As String)
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim s As BaseStore = New CustomStore()
        __Check(CStr(s(10)), "CustomStore_10")
    End Sub
End Module
