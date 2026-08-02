' vybe-test: vb/vb_legacy_math_functions/legacy_math_functions
' origin: languages/vb/tests/vb/test_vb_legacy_math_functions.rs

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
        ' Initialize random-number generator
        Randomize(42)
        
        ' Rnd returns a Single less than 1 but greater than or equal to 0
        Dim val1 = Rnd()
        __Check(CStr(val1 >= 0 AndAlso val1 < 1), "True")
        
        ' Int returns the integer portion of a number
        __Check(CStr(Int(12.34)), "12")
        __Check(CStr(Int(-12.34)), "-13") ' Int rounds down (-13)
        
        ' Fix returns the integer portion of a number
        __Check(CStr(Fix(12.34)), "12")
        __Check(CStr(Fix(-12.34)), "-12") ' Fix truncates towards zero (-12)
    End Sub
End Module
