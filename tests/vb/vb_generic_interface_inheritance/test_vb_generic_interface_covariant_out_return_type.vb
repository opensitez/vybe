' vybe-test: vb/vb_generic_interface_inheritance/test_vb_generic_interface_covariant_out_return_type
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

Interface IProvider(Of Out T)
    Function GetItem() As T
End Interface

Class StringProvider
    Implements IProvider(Of String)
    Public Function GetItem() As String Implements IProvider(Of String).GetItem
        Return "Provided String"
    End Function
End Class

Module Program
    Sub Main()
        Dim provider As IProvider(Of Object) = New StringProvider()
        __Check(CStr(provider.GetItem().ToString()), "Provided String")
    End Sub
End Module
