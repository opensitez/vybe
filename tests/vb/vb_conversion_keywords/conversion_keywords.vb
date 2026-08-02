' vybe-test: vb/vb_conversion_keywords/conversion_keywords
' origin: languages/vb/tests/vb/test_vb_conversion_keywords.rs

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
        Dim s As String = "123"
        Dim d As Double = 45.67
        
        ' CInt converts to Integer (rounds)
        Dim i As Integer = CInt(d)
        __Check(CStr(i), "46") ' Rounds 45.67 to 46
        
        ' CStr converts to String
        Dim strVal = CStr(100)
        __Check(CStr(strVal), "100")
        
        ' CDbl converts to Double
        Dim dVal = CDbl(s)
        __Check(CStr(dVal + 1), "124")
        
        ' CBool converts to Boolean
        Dim bVal = CBool("True")
        __Check(CStr(bVal), "True")
    End Sub
End Module
