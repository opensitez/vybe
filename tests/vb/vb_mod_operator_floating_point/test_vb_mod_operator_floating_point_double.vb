' vybe-test: vb/vb_mod_operator_floating_point/test_vb_mod_operator_floating_point_double
' origin: languages/vb/tests/vb/test_vb_mod_operator_floating_point.rs

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

Module Program
    Sub Main()
        ' Unlike C#, VB.NET Mod rounds floating point operands to Long before computing Mod!
        Dim a As Double = 17.6 ' Rounds to 18
        Dim b As Double = 4.9  ' Rounds to 5
        Dim res = a Mod b      ' 18 Mod 5 = 3
        __Check(CStr(res), "3")
    End Sub
End Module
