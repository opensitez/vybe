' vybe-test: vb/vb_complex_class_hierarchy_generics/test_vb_generic_contravariance_in_parameter
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

Interface IReceiver(Of In T)
    Sub Receive(data As T)
End Interface

Class ObjectReceiver
    Implements IReceiver(Of Object)
    Public Sub Receive(data As Object) Implements IReceiver(Of Object).Receive
        __Check(CStr("Received: " & data.ToString()), "Received: ContravariantPayload")
    End Sub
End Class

Module Program
    Sub Main()
        Dim objRec As IReceiver(Of Object) = New ObjectReceiver()
        Dim strRec As IReceiver(Of String) = objRec ' Contravariant assignment!
        strRec.Receive("ContravariantPayload")
    End Sub
End Module
