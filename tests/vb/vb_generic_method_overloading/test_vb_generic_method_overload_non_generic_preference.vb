' vybe-test: vb/vb_generic_method_overloading/test_vb_generic_method_overload_non_generic_preference
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

Module Utility
    Public Sub Display(Of T)(val As T)
        __Check(CStr("Generic: " & val.ToString()), "NonGenericString: Hello")
    End Sub

    Public Sub Display(val As String)
        __Check(CStr("NonGenericString: " & val), "Generic: 123")
    End Sub
End Module

Module Program
    Sub Main()
        Utility.Display("Hello")
        Utility.Display(123)
    End Sub
End Module
