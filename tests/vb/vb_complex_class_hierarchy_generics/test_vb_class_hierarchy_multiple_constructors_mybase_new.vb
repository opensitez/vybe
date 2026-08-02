' vybe-test: vb/vb_complex_class_hierarchy_generics/test_vb_class_hierarchy_multiple_constructors_mybase_new
' origin: languages/vb/tests/vb/test_vb_complex_class_hierarchy_generics.rs

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

Class BasePerson
    Public Name As String
    Public Sub New(n As String)
        Name = n
    End Sub
End Class

Class Employee
    Inherits BasePerson

    Public Salary As Decimal
    Public Sub New(n As String, s As Decimal)
        MyBase.New(n)
        Salary = s
    End Sub
End Class

Module Program
    Sub Main()
        Dim emp As New Employee("Alice", 75000D)
        __Check(CStr(emp.Name & ":" & emp.Salary), "Alice:75000")
    End Sub
End Module
