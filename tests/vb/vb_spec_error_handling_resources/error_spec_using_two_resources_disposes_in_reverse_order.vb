' vybe-test: vb/vb_spec_error_handling_resources/error_spec_using_two_resources_disposes_in_reverse_order
' origin: languages/vb/tests/vb/test_vb_spec_error_handling_resources.rs

Class Probe
    Implements IDisposable
    Private _name As String
    Public Sub New(name As String)
        _name = name
    End Sub
    Public Sub Dispose() Implements IDisposable.Dispose
        Console.WriteLine(_name)
    End Sub
End Class
Module M
    Sub Main()
        Using first As New Probe("first"), second As New Probe("second")
            Console.WriteLine("body")
        End Using
    End Sub
End Module
