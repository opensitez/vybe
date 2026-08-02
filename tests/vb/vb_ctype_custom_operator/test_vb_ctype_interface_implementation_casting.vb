' vybe-test: vb/vb_ctype_custom_operator/test_vb_ctype_interface_implementation_casting
' origin: languages/vb/tests/vb/test_vb_ctype_custom_operator.rs

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

Interface IStorable
    Sub Save()
End Interface

Class Document
    Implements IStorable
    Public Sub Save() Implements IStorable.Save
        __Check(CStr("Document Saved"), "Document Saved")
    End Sub
End Class

Module Program
    Sub Main()
        Dim doc As New Document()
        Dim storable As IStorable = CType(doc, IStorable)
        storable.Save()
    End Sub
End Module
