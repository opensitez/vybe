' vybe-test: vb/vb_trycast_reference_types/test_vb_trycast_same_type_identity_cast
' origin: languages/vb/tests/vb/test_vb_trycast_reference_types.rs

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

Class Node
End Class

Module Program
    Sub Main()
        Dim n As New Node()
        Dim n2 As Node = TryCast(n, Node)
        __Check(CStr(Object.ReferenceEquals(n, n2)), "True")
    End Sub
End Module
