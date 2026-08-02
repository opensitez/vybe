' vybe-test: vb/vb_class_nested_private_public/test_vb_nested_generic_class_inside_non_generic_class
' origin: languages/vb/tests/vb/test_vb_class_nested_private_public.rs

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

Class Storage
    Public Class Cache(Of T)
        Private item As T
        Public Sub SetItem(val As T) : item = val : End Sub
        Public Function GetItem() As T : Return item : End Function
    End Class
End Class

Module Program
    Sub Main()
        Dim c As New Storage.Cache(Of String)()
        c.SetItem("CachedData")
        __Check(CStr(c.GetItem()), "CachedData")
    End Sub
End Module
