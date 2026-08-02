' vybe-test: vb/vb_byref_mutation/byref_can_rebind_reference_type_variable_in_caller
' origin: languages/vb/tests/vb/test_vb_byref_mutation.rs

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

Module M
    Class Holder
        Public Value As String
        Public Sub New(value As String)
            Me.Value = value
        End Sub
    End Class

    Sub Replace(ByRef holder As Holder)
        holder = New Holder("replaced")
    End Sub

    Sub Main()
        Dim value As Holder = New Holder("initial")
        Replace(value)
        __Check(CStr(value.Value), "replaced")
    End Sub
End Module
