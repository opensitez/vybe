' vybe-test: vb/vb_spec_error_handling_resources/error_spec_synclock_inside_using_can_share_same_object
' origin: languages/vb/tests/vb/test_vb_spec_error_handling_resources.rs

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

Class Holder
    Implements IDisposable
    Public Gate As New Object()
    Public Sub Dispose() Implements IDisposable.Dispose
        __Check(CStr("disposed"), "locked")
    End Sub
End Class
Module M
    Sub Main()
        Using holder As New Holder()
            SyncLock holder.Gate
                __Check(CStr("locked"), "disposed")
            End SyncLock
        End Using
    End Sub
End Module
