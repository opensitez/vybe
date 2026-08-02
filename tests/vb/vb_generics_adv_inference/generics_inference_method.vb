' vybe-test: vb/vb_generics_adv_inference/generics_inference_method
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
    Function CreateArray(Of T)(item1 As T, item2 As T) As T()
        Return {item1, item2}
    End Function

    Sub Main()
        ' Type is inferred from arguments
        Dim arr = CreateArray(10, 20)
        __Check(CStr(arr(0)), "10")
        __Check(CStr(arr(1)), "20")
    End Sub
End Module
