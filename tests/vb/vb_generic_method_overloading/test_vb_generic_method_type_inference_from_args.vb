' vybe-test: vb/vb_generic_method_overloading/test_vb_generic_method_type_inference_from_args
' origin: languages/vb/tests/vb/test_vb_generic_method_overloading.rs

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

Module Helper
    Public Function Identity(Of T)(item As T) As T
        Return item
    End Function
End Module

Module Program
    Sub Main()
        Dim resStr = Helper.Identity("InferString")
        Dim resInt = Helper.Identity(42)
        __Check(CStr(resStr), "InferString")
        __Check(CStr(resInt), "42")
    End Sub
End Module
