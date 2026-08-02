' vybe-test: vb/vb_goto_label_forward_backward/test_vb_goto_out_of_try_block
' origin: languages/vb/tests/vb/test_vb_goto_label_forward_backward.rs

Imports System

Module Program
    Sub Main()
        Try
            Console.WriteLine("In Try Block")
            GoTo ExternalLabel
        Catch ex As Exception
            Console.WriteLine("In Catch")
        Finally
            Console.WriteLine("In Finally")
        End Try
ExternalLabel:
        Console.WriteLine("Outside Try")
    End Sub
End Module
