' vybe-test: vb/vb_class/class_with_constructor
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
    Class Person
        Public Name As String
        Public Age As Integer

        Sub New(n As String, a As Integer)
            Me.Name = n
            Me.Age = a
        End Sub
    End Class

    Sub Main()
        Dim p As New Person("Bob", 25)
        __Check(CStr(p.Name & " is " & CStr(p.Age)), "Bob is 25")
    End Sub
End Module
