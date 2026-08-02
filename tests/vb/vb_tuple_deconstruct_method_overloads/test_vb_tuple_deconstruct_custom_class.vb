' vybe-test: vb/vb_tuple_deconstruct_method_overloads/test_vb_tuple_deconstruct_custom_class
' origin: languages/vb/tests/vb/test_vb_tuple_deconstruct_method_overloads.rs

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

Class Person
    Public Property Name As String
    Public Property Age As Integer
    Public Sub New(n As String, a As Integer) : Name = n : Age = a : End Sub
    Public Sub Deconstruct(ByRef n As String, ByRef a As Integer)
        n = Name : a = Age
    End Sub
End Class

Module Program
    Sub Main()
        Dim p As New Person("Bob", 40)
        Dim n As String = Nothing
        Dim a As Integer = 0
        p.Deconstruct(n, a)
        __Check(CStr(n & " is " & a), "Bob is 40")
    End Sub
End Module
