' vybe-test: vb/vb_generic_constraints_multiple/test_vb_generic_constraint_class_and_new
' origin: languages/vb/tests/vb/test_vb_generic_constraints_multiple.rs

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

Class Factory(Of T As {Class, New})
    Public Function CreateInstance() As T
        Return New T()
    End Function
End Class

Class Item
    Public Tag As String = "Created"
End Class

Module Program
    Sub Main()
        Dim f As New Factory(Of Item)()
        Dim i As Item = f.CreateInstance()
        __Check(CStr(i.Tag), "Created")
    End Sub
End Module
