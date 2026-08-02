' vybe-test: vb/vb_lset_rset/lset_rset_statements
' origin: languages/vb/tests/vb/test_vb_lset_rset.rs

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
        ' LSet and RSet pad strings with spaces to match the length of the target variable
        Dim s1 As String = "1234567890"
        LSet s1 = "Left"
        __Check(CStr("[" & s1 & "]"), "[Left      ]")
        
        Dim s2 As String = "1234567890"
        RSet s2 = "Right"
        __Check(CStr("[" & s2 & "]"), "[     Right]")
    End Sub
End Module
