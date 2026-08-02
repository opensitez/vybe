' vybe-test: vb/vb_class_nested_private_public/test_vb_nested_delegate_declaration
' origin: languages/vb/tests/vb/test_vb_class_nested_private_public.rs

Class Processor
    Public Delegate Sub ProgressHandler(percent As Integer)

    Public Sub Run(handler As ProgressHandler)
        handler(50)
        handler(100)
    End Sub
End Class

Module Program
    Sub Main()
        Dim p As New Processor()
        p.Run(Sub(pct) Console.WriteLine("Progress: " & pct & "%"))
    End Sub
End Module
