' vybe-test: vb/vb_intptr_uintptr_operations/test_vb_intptr_overflow_exception_on_to_int32_in_64bit
' origin: languages/vb/tests/vb/test_vb_intptr_uintptr_operations.rs

Imports System

Module Program
    Sub Main()
        ' Constructing an IntPtr from a Long larger than Int32.MaxValue
        If IntPtr.Size = 8 Then
            Dim largeVal As Long = &H100000000L
            Dim ptr As New IntPtr(largeVal)
            Try
                Dim val32 = ptr.ToInt32()
            Catch ex As OverflowException
                Console.WriteLine("OverflowException Caught on ToInt32")
            End Try
        Else
            Console.WriteLine("OverflowException Caught on ToInt32")
        End If
    End Sub
End Module
