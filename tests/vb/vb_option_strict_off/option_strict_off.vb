' vybe-test: vb/vb_option_strict_off/option_strict_off
' origin: languages/vb/tests/vb/test_vb_option_strict_off.rs

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

Module M
    Sub Main()
        ' With Option Strict Off, implicit narrowing conversions are allowed
        Dim x As Double = 42.5
        Dim y As Integer = x ' Implicit conversion to Integer, rounds to 42
        __Check(CStr(y), "42")
        
        ' And Late Binding is allowed
        Dim obj As Object = "Hello"
        __Check(CStr(obj.ToUpper()), "HELLO")
    End Sub
End Module
