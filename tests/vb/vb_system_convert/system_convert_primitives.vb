' vybe-test: vb/vb_system_convert/system_convert_primitives
' origin: languages/vb/tests/vb/test_vb_system_convert.rs

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

Imports System

Module M
    Sub Main()
        Dim s As String = "123"
        Dim i As Integer = Convert.ToInt32(s)
        __Check(CStr(i), "123")
        
        Dim b As Boolean = Convert.ToBoolean(1)
        __Check(CStr(b), "True")
        
        Dim d As Double = Convert.ToDouble("3.14", Globalization.CultureInfo.InvariantCulture)
        __Check(CStr(d), "3.14")
    End Sub
End Module
