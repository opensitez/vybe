' vybe-test: vb/vb_property_indexed_multi_parameter/test_vb_indexed_property_paramarray_arguments
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

Class VariadicIndexer
    Public Property KeyPath(ParamArray keys As String()) As String
        Get
            Return String.Join("/", keys)
        End Get
        Set(value As String)
        End Set
    End Property
End Class

Module Program
    Sub Main()
        Dim v As New VariadicIndexer()
        __Check(CStr(v.KeyPath("usr", "local", "bin")), "usr/local/bin")
    End Sub
End Module
