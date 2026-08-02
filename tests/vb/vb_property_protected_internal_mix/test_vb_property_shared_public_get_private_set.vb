' vybe-test: vb/vb_property_protected_internal_mix/test_vb_property_shared_public_get_private_set
' origin: languages/vb/tests/vb/test_vb_property_protected_internal_mix.rs

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

Class GlobalCounter
    Public Shared Property TotalCount As Integer { Get; Private Set; } = 0
    Public Shared Sub Increment()
        TotalCount += 1
    End Sub
End Class

Module Program
    Sub Main()
        GlobalCounter.Increment()
        GlobalCounter.Increment()
        __Check(CStr(GlobalCounter.TotalCount), "2")
    End Sub
End Module
