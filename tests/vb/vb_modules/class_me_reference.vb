' vybe-test: vb/vb_modules/class_me_reference
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

Class Counter
    Public Value As Integer = 0
    Public Sub Increment()
        Me.Value = Me.Value + 1
    End Sub
End Class

Module M
    Sub Main()
        Dim c As New Counter()
        c.Increment()
        c.Increment()
        __Check(CStr(c.Value), "2")
    End Sub
End Module
