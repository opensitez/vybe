' vybe-test: vb/vb_interface_default_methods_adv/test_vb_interface_co_variance_out_keyword
' origin: languages/vb/tests/vb/test_vb_interface_default_methods_adv.rs

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

Interface IProducer(Of Out T)
    Function Produce() As T
End Interface

Class StringProducer
    Implements IProducer(Of String)
    Public Function Produce() As String Implements IProducer(Of String).Produce
        Return "Produced String"
    End Function
End Class

Module Program
    Sub Main()
        Dim p As IProducer(Of Object) = New StringProducer()
        __Check(CStr(p.Produce().ToString()), "Produced String")
    End Sub
End Module
