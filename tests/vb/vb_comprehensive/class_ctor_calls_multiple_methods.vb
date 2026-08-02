' vybe-test: vb/vb_comprehensive/class_ctor_calls_multiple_methods
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
    Class Setup
        Public A As String
        Public B As String

        Sub New()
            SetupA()
            SetupB()
        End Sub

        Sub SetupA()
            Me.A = "alpha"
        End Sub

        Sub SetupB()
            Me.B = "beta"
        End Sub
    End Class

    Sub Main()
        Dim s As New Setup()
        __Check(CStr(s.A), "alpha")
        __Check(CStr(s.B), "beta")
    End Sub
End Module
