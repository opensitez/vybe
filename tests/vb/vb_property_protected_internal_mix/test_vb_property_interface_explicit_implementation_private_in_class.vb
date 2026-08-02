' vybe-test: vb/vb_property_protected_internal_mix/test_vb_property_interface_explicit_implementation_private_in_class
' origin: languages/vb/tests/vb/test_vb_property_protected_internal_mix.rs

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

Interface IInternalData
    ReadOnly Property Data As String
End Interface

Class SecretProvider
    Implements IInternalData
    Private ReadOnly Property Data As String Implements IInternalData.Data
        Get
            Return "SecretDataValue"
        End Get
    End Property
End Class

Module Program
    Sub Main()
        Dim provider As IInternalData = New SecretProvider()
        __Check(CStr(provider.Data), "SecretDataValue")
    End Sub
End Module
