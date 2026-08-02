' vybe-test: vb/vb_comprehensive/class_inheritance
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
    Class Animal
        Public Name As String

        Sub New(n As String)
            Me.Name = n
        End Sub

        Function Describe() As String
            Describe = "Animal: " & Me.Name
        End Function
    End Class

    Class Dog
        Inherits Animal

        Sub New(n As String)
            MyBase.New(n)
        End Sub

        Function Bark() As String
            Bark = Me.Name & " barks!"
        End Function
    End Class

    Sub Main()
        Dim d As New Dog("Rex")
        __Check(CStr(d.Describe()), "Animal: Rex")
        __Check(CStr(d.Bark()), "Rex barks!")
    End Sub
End Module
