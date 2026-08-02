' vybe-test: vb/vb_generic_interface_inheritance/test_vb_generic_interface_contravariant_in_parameter
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

Interface IConsumer(Of In T)
    Sub Consume(item As T)
End Interface

Class ObjectConsumer
    Implements IConsumer(Of Object)
    Public Sub Consume(item As Object) Implements IConsumer(Of Object).Consume
        __Check(CStr("Consuming: " & item.ToString()), "Consuming: Test Message")
    End Sub
End Class

Module Program
    Sub Main()
        Dim c As IConsumer(Of String) = New ObjectConsumer()
        c.Consume("Test Message")
    End Sub
End Module
