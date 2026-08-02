' vybe-test: vb/vb_interfaces_multiple/interface_implementing_same_method_from_multiple
' origin: languages/vb/tests/vb/test_vb_interfaces_multiple.rs

Interface IControl
    Sub Paint()
End Interface

Interface ISurface
    Sub Paint()
End Interface

Class Canvas
    Implements IControl, ISurface
    
    ' One method satisfies both interfaces
    Public Sub Paint() Implements IControl.Paint, ISurface.Paint
        Console.WriteLine("Painted both")
    End Sub
End Class

Module M
    Sub Main()
        Dim c As New Canvas()
        Dim c1 As IControl = c
        Dim c2 As ISurface = c
        
        c1.Paint()
        c2.Paint()
    End Sub
End Module
