' vybe-test: vb/vb_spec_error_handling_resources/error_spec_nested_using_scopes_dispose_both_resources
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
        Using outerValue As New Probe("outer")
            Using innerValue As New Probe("inner")
                Console.WriteLine("body")
            End Using
        End Using
    End Sub
End Module
