' vybe-test: vb/vb_static_properties/static_properties
' origin: languages/vb/tests/vb/test_vb_static_properties.rs

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

Class Cache
    ' Static properties maintain state across all instances
    Public Shared Property LastAccessed As String = "None"
    Public Shared ReadOnly Property CreatedAt As Date = #2024-01-01#
    
    Public Sub Access(item As String)
        LastAccessed = item
    End Sub
End Class

Module M
    Sub Main()
        Dim c1 As New Cache()
        c1.Access("Item1")
        
        Dim c2 As New Cache()
        __Check(CStr(Cache.LastAccessed), "Item1")
        
        c2.Access("Item2")
        __Check(CStr(Cache.LastAccessed), "Item2")
        
        __Check(CStr(Cache.CreatedAt.Year), "2024")
    End Sub
End Module
