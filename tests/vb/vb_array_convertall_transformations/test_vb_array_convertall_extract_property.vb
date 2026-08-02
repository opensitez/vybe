' vybe-test: vb/vb_array_convertall_transformations/test_vb_array_convertall_extract_property
' origin: languages/vb/tests/vb/test_vb_array_convertall_transformations.rs

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

Imports System

Class Product
    Public Property Id As Integer
    Public Sub New(id As Integer)
        Me.Id = id
    End Sub
End Class

Module Program
    Sub Main()
        Dim prods As Product() = {New Product(101), New Product(102)}
        Dim ids As Integer() = Array.ConvertAll(prods, Function(p) p.Id)
        __Check(CStr(String.Join("-", ids)), "101-102")
    End Sub
End Module
