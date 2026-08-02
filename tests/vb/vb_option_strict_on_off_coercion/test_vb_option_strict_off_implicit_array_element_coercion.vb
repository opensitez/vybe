' vybe-test: vb/vb_option_strict_on_off_coercion/test_vb_option_strict_off_implicit_array_element_coercion
' origin: languages/vb/tests/vb/test_vb_option_strict_on_off_coercion.rs

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

Option Strict Off

Module Program
    Sub Main()
        Dim rawObjs As Object() = New Object() {"10", "20", "30"}
        Dim firstNum As Integer = rawObjs(0)
        __Check(CStr(firstNum * 2), "20")
    End Sub
End Module
