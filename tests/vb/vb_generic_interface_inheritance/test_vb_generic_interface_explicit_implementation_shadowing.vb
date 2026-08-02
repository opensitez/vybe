' vybe-test: vb/vb_generic_interface_inheritance/test_vb_generic_interface_explicit_implementation_shadowing
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

Interface IProcessor(Of T)
    Sub Process(data As T)
End Interface

Class Processor
    Implements IProcessor(Of String)
    Private Sub Process(data As String) Implements IProcessor(Of String).Process
        __Check(CStr("Explicit Process: " & data), "Explicit Process: Data")
    End Sub
End Class

Module Program
    Sub Main()
        Dim p As IProcessor(Of String) = New Processor()
        p.Process("Data")
    End Sub
End Module
