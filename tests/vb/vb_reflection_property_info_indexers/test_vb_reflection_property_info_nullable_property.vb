' vybe-test: vb/vb_reflection_property_info_indexers/test_vb_reflection_property_info_nullable_property
' origin: languages/vb/tests/vb/test_vb_reflection_property_info_indexers.rs

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

Class Document
    Public Property Pages As Nullable(Of Integer) = 42
End Class

Module Program
    Sub Main()
        Dim doc As New Document()
        Dim prop = GetType(Document).GetProperty("Pages")
        Dim val = prop.GetValue(doc)
        __Check(CStr(val.GetType().Name & "=" & val), "Int32=42")
    End Sub
End Module
