' vybe-test: vb/vb_generic_static_member_per_type/test_vb_generic_shared_constructor_per_type_counter
' origin: languages/vb/tests/vb/test_vb_generic_static_member_per_type.rs

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

Class Tracker(Of T)
    Public Shared TotalInits As Integer = 0
    Shared Sub New()
        TotalInits += 1
    End Sub
End Class

Module Program
    Sub Main()
        ' Accessing Integer and Double triggers two distinct Shared Sub New invocations
        __Check(CStr(Tracker(Of Integer).TotalInits & "|" & Tracker(Of Double).TotalInits), "1|1")
    End Sub
End Module
