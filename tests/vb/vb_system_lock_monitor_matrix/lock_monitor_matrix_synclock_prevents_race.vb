' vybe-test: vb/vb_system_lock_monitor_matrix/lock_monitor_matrix_synclock_prevents_race
' origin: languages/vb/tests/vb/test_vb_system_lock_monitor_matrix.rs

Imports System
Imports System.Threading

Module M
    Sub Main()
        Dim lockObject As New Object()
        Dim value As Integer = 0

        Dim t1 As New Thread(
            Sub()
                For i As Integer = 0 To 999
                    SyncLock lockObject
                        value += 1
                    End SyncLock
                Next
            End Sub)
        Dim t2 As New Thread(
            Sub()
                For i As Integer = 0 To 999
                    SyncLock lockObject
                        value += 1
                    End SyncLock
                Next
            End Sub)

        t1.Start()
        t2.Start()
        t1.Join()
        t2.Join()

        Console.WriteLine(value)
    End Sub
End Module
