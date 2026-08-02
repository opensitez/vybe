' vybe-test: vb/vb_comprehensive/multiple_class_instances
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
    Class Dog
        Public Name As String

        Sub New(n As String)
            Me.Name = n
        End Sub

        Function Speak() As String
            Speak = Me.Name & " says Woof!"
        End Function
    End Class

    Sub Main()
        Dim a As New Dog("Rex")
        Dim b As New Dog("Buddy")
        __Check(CStr(a.Speak()), "Rex says Woof!")
        __Check(CStr(b.Speak()), "Buddy says Woof!")
    End Sub
End Module
