' vybe-test: vb/vb_math_trig_log/math_trig_log
' origin: languages/vb/tests/vb/test_vb_math_trig_log.rs

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
        ' Abs
        __Check(CStr(Abs(-15.5)), "15.5")
        
        ' Sqrt
        __Check(CStr(Sqrt(16)), "4")
        
        ' Trig
        __Check(CStr(Int(Cos(0))), "1")
        __Check(CStr(Int(Sin(0))), "0")
        __Check(CStr(Int(Tan(0))), "0")
        
        ' Log / Exp
        __Check(CStr(Exp(0)), "1")
        __Check(CStr(Int(Log(Exp(1)))), "1")
    End Sub
End Module
