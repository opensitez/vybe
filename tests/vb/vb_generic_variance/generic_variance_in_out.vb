' vybe-test: vb/vb_generic_variance/generic_variance_in_out
' origin: languages/vb/tests/vb/test_vb_generic_variance.rs

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

' Out defines covariance, In defines contravariance
Interface IProducer(Of Out T)
    Function Produce() As T
End Interface

Interface IConsumer(Of In T)
    Sub Consume(item As T)
End Interface

Class StringProducer
    Implements IProducer(Of String)
    
    Public Function Produce() As String Implements IProducer(Of String).Produce
        Return "Hello"
    End Function
End Class

Module M
    Sub Main()
        ' Covariance: Assigning a more specific generic type to a less specific one
        Dim p As IProducer(Of Object) = New StringProducer()
        __Check(CStr(p.Produce()), "Hello")
    End Sub
End Module
