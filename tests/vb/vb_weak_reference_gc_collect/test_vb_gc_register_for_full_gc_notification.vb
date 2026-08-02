' vybe-test: vb/vb_weak_reference_gc_collect/test_vb_gc_register_for_full_gc_notification
' origin: languages/vb/tests/vb/test_vb_weak_reference_gc_collect.rs

Imports System

Module Program
    Sub Main()
        Try
            GC.RegisterForFullGCNotification(10, 10)
            GC.CancelFullGCNotification()
            Console.WriteLine("GC Notification Handled")
        Catch ex As Exception
            Console.WriteLine("GC Notification Unsupported or Ok")
        End Try
    End Sub
End Module
