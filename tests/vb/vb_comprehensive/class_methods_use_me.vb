' vybe-test: vb/vb_comprehensive/class_methods_use_me
' origin: languages/vb/tests/vb/vb_comprehensive_test.rs

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
    Class Box
        Public Width As Integer
        Public Height As Integer

        Sub New(w As Integer, h As Integer)
            Me.Width = w
            Me.Height = h
        End Sub

        Function Area() As Integer
            Area = Me.Width * Me.Height
        End Function
    End Class

    Sub Main()
        Dim b As New Box(5, 3)
        __Check(CStr(b.Area()), "15")
    End Sub
End Module
