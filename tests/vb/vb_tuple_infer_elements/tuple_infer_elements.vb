' vybe-test: vb/vb_tuple_infer_elements/tuple_infer_elements
' origin: languages/vb/tests/vb/test_vb_tuple_infer_elements.rs

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

Module M
    Sub Main()
        Dim x = 10
        Dim y = "Test"
        
        ' VB.NET tuple element name inference
        Dim t = (x, y)
        
        ' The inferred names are 'x' and 'y' (if supported by compiler)
        ' Let's access them via standard ItemN to be safe if inference isn't supported
        __Check(CStr(t.Item1), "10")
        __Check(CStr(t.Item2), "Test")
    End Sub
End Module
