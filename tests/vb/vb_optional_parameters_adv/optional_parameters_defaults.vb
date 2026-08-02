' vybe-test: vb/vb_optional_parameters_adv/optional_parameters_defaults
' origin: languages/vb/tests/vb/test_vb_optional_parameters_adv.rs

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
    Function Greet(name As String, Optional greeting As String = "Hello", Optional punctuation As String = "!") As String
        Return greeting & " " & name & punctuation
    End Function

    Sub Main()
        __Check(CStr(Greet("Alice")), "Hello Alice!")
        __Check(CStr(Greet("Bob", "Hi")), "Hi Bob!")
        __Check(CStr(Greet("Charlie", "Hey", "?")), "Hey Charlie?")
        
        ' Named parameters skipping optional
        __Check(CStr(Greet("Dave", punctuation:=".")), "Hello Dave.")
    End Sub
End Module
