' vybe-test: vb/vb_synclock/synclock_basic
' origin: languages/vb/tests/vb/test_vb_synclock.rs

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

Class Resource
    Public Value As Integer = 0
End Class

Module M
    Private _lockObj As New Object()
    
    Sub Main()
        Dim res As New Resource()
        
        ' Note: actual multithreading test is complex in basic VM output,
        ' but we can test that the SyncLock syntax parses and executes the block.
        SyncLock _lockObj
            res.Value = res.Value + 10
            __Check(CStr("Locked and loaded"), "Locked and loaded")
        End SyncLock
        
        __Check(CStr(res.Value), "10")
    End Sub
End Module
