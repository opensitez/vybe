' vybe-test: vb/vb_partial_method_byref/partial_method_byref
' origin: languages/vb/tests/vb/test_vb_partial_method_byref.rs

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

Partial Class Processor
    Partial Private Sub ModifyValue(ByRef val As Integer)
    End Sub
    
    Public Sub Run()
        Dim x = 10
        ModifyValue(x)
        __Check(CStr(x), "20")
    End Sub
End Class

Partial Class Processor
    Private Sub ModifyValue(ByRef val As Integer)
        val = 20
    End Sub
End Class

Module M
    Sub Main()
        Dim p As New Processor()
        p.Run()
    End Sub
End Module
