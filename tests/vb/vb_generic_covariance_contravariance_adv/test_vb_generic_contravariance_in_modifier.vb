' vybe-test: vb/vb_generic_covariance_contravariance_adv/test_vb_generic_contravariance_in_modifier
' origin: languages/vb/tests/vb/test_vb_generic_covariance_contravariance_adv.rs

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

Public Interface IItemConsumer(Of In T)
    Sub Consume(item As T)
End Interface

Class ObjectConsumer
    Implements IItemConsumer(Of Object)
    Public Sub Consume(item As Object) Implements IItemConsumer(Of Object).Consume
        __Check(CStr("Consumed: " & item.ToString()), "Consumed: ContravariantValue")
    End Sub
End Class

Module Program
    Sub Main()
        Dim objCons As IItemConsumer(Of Object) = New ObjectConsumer()
        Dim strCons As IItemConsumer(Of String) = objCons ' Contravariance assignment
        strCons.Consume("ContravariantValue")
    End Sub
End Module
