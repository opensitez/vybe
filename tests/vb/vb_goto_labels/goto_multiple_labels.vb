' vybe-test: vb/vb_goto_labels/goto_multiple_labels
' origin: languages/vb/tests/vb/test_vb_goto_labels.rs

Module M
    Sub Main()
        Dim x As Integer = 2
        If x = 1 Then GoTo Label1
        If x = 2 Then GoTo Label2
        If x = 3 Then GoTo Label3
        
Label1:
        Console.WriteLine("L1")
        Exit Sub
Label2:
        Console.WriteLine("L2")
        Exit Sub
Label3:
        Console.WriteLine("L3")
    End Sub
End Module
