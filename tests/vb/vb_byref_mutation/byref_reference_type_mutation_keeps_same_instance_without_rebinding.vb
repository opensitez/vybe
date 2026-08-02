' vybe-test: vb/vb_byref_mutation/byref_reference_type_mutation_keeps_same_instance_without_rebinding
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
        Public Value As Integer
        Public Sub New(value As Integer)
            Me.Value = value
        End Sub
    End Class

    Sub Boost(ByRef holder As Holder)
        holder.Value = holder.Value + 3
    End Sub

    Sub Main()
        Dim value As Holder = New Holder(7)
        Boost(value)
        __Check(CStr(value.Value), "10")
    End Sub
End Module
