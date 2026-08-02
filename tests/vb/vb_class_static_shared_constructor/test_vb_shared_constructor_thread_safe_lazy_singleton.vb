' vybe-test: vb/vb_class_static_shared_constructor/test_vb_shared_constructor_thread_safe_lazy_singleton
' origin: languages/vb/tests/vb/test_vb_class_static_shared_constructor.rs

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

Class LazySingleton
    Public Shared ReadOnly Instance As LazySingleton
    Public Property CreatedAt As String
    Shared Sub New()
        Instance = New LazySingleton() With {.CreatedAt = "CreatedInSharedSubNew"}
    End Sub
    Private Sub New()
    End Sub
End Class

Module Program
    Sub Main()
        __Check(CStr(LazySingleton.Instance.CreatedAt), "CreatedInSharedSubNew")
    End Sub
End Module
