' vybe-test: vb/vb_modules/class_inheritance
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

Class Animal
    Public Name As String
    Public Sub New(n As String)
        Name = n
    End Sub
    Public Function Speak() As String
        Return Name & " makes a sound"
    End Function
End Class

Class Dog
    Inherits Animal
    Public Sub New(n As String)
        MyBase.New(n)
    End Sub
    Public Overrides Function Speak() As String
        Return Name & " says Woof!"
    End Function
End Class

Module M
    Sub Main()
        Dim d As New Dog("Rex")
        __Check(CStr(d.Speak()), "Rex says Woof!")
    End Sub
End Module
