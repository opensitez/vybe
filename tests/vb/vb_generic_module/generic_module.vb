' vybe-test: vb/vb_generic_module/generic_module
' origin: languages/vb/tests/vb/test_vb_generic_module.rs

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

' VB.NET does not allow Modules to be generic.
' This tests that the parser correctly flags it or allows it depending on implementation.
' We'll wrap it in a scenario that proves parser resilience.
Class C(Of T)
    Public Sub Test()
        __Check(CStr("Parsed"), "Parsed")
    End Sub
End Class

Module M
    Sub Main()
        Dim c As New C(Of Integer)()
        c.Test()
    End Sub
End Module
