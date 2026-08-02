' vybe-test: vb/vb_generics_adv_inference/generics_inference_lambda
' origin: languages/vb/tests/vb/test_vb_generics_adv_inference.rs

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
        Dim process = Function(Of T)(val As T) val
        
        __Check(CStr(process("Hello")), "Hello")
        __Check(CStr(process(42)), "42")
    End Sub
End Module
