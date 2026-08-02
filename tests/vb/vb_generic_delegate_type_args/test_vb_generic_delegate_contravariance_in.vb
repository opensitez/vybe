' vybe-test: vb/vb_generic_delegate_type_args/test_vb_generic_delegate_contravariance_in
' origin: languages/vb/tests/vb/test_vb_generic_delegate_type_args.rs

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

Delegate Sub Consumer(Of In T)(item As T)

Module Program
    Private Sub ConsumeObject(obj As Object)
        __Check(CStr("Contravariant: " & obj.ToString()), "Contravariant: InputString")
    End Sub

    Sub Main()
        Dim c As Consumer(Of String) = AddressOf ConsumeObject
        c("InputString")
    End Sub
End Module
