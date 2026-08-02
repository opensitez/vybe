' vybe-test: vb/vb_system_boxing_unboxing_matrix/boxing_unboxing_reference_preservation_on_classes
' origin: languages/vb/tests/vb/test_vb_system_boxing_unboxing_matrix.rs

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

Class Holder
    Public Value As Integer

    Public Sub New(v As Integer)
        Value = v
    End Sub
End Class

Module M
    Sub Main()
        Dim h As New Holder(9)
        Dim boxed As Object = h
        Dim unboxed As Holder = CType(boxed, Holder)

        __Check(CStr(unboxed.Value), "9")
        unboxed.Value = 10
        __Check(CStr(h.Value), "10")
    End Sub
End Module
