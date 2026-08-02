' vybe-test: vb/vb_modules/pass_object_to_sub
' origin: languages/vb/tests/vb/test_vb_modules.rs

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

Class Box
    Public Value As Integer
End Class

Module M
    Sub SetValue(b As Box, v As Integer)
        b.Value = v
    End Sub
    Sub Main()
        Dim b As New Box()
        SetValue(b, 42)
        __Check(CStr(b.Value), "42")
    End Sub
End Module
