' vybe-test: vb/vb_class/class_with_field_initializer
' origin: languages/vb/tests/vb/vb_class_test.rs

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
    Class Counter
        Public Count As Integer = 0

        Sub Increment()
            Me.Count = Me.Count + 1
        End Sub
    End Class

    Sub Main()
        Dim c As New Counter()
        c.Increment()
        c.Increment()
        c.Increment()
        __Check(CStr(c.Count), "3")
    End Sub
End Module
