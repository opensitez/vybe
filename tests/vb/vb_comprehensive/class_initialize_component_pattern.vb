' vybe-test: vb/vb_comprehensive/class_initialize_component_pattern
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
    Class MyForm
        Public Title As String

        Sub New()
            InitializeComponent()
        End Sub

        Sub InitializeComponent()
            Me.Title = "My Application"
        End Sub
    End Class

    Sub Main()
        Dim f As New MyForm()
        __Check(CStr(f.Title), "My Application")
    End Sub
End Module
