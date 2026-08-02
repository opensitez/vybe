' vybe-test: vb/vb_synclock_advanced/synclock_advanced
' origin: languages/vb/tests/vb/test_vb_synclock_advanced.rs

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

Imports System.Threading

Class Counter
    Private _count As Integer = 0
    Private _lockObj As New Object()
    
    Public Sub Increment()
        SyncLock _lockObj
            _count += 1
        End SyncLock
    End Sub
    
    Public ReadOnly Property Count As Integer
        Get
            SyncLock _lockObj
                Return _count
            End SyncLock
        End Get
    End Property
End Class

Module M
    Sub Main()
        Dim c As New Counter()
        
        Dim t1 As New Thread(Sub() c.Increment())
        Dim t2 As New Thread(Sub() c.Increment())
        
        t1.Start()
        t2.Start()
        
        t1.Join()
        t2.Join()
        
        __Check(CStr(c.Count), "2")
    End Sub
End Module
