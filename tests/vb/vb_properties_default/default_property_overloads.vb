' vybe-test: vb/vb_properties_default/default_property_overloads
' origin: languages/vb/tests/vb/test_vb_properties_default.rs

Class DictionaryMock
    Private _keys(10) As String
    Private _values(10) As String
    Private _count As Integer = 0
    
    Public Sub Add(key As String, value As String)
        _keys(_count) = key
        _values(_count) = value
        _count += 1
    End Sub

    Default Public ReadOnly Property Item(key As String) As String
        Get
            For i As Integer = 0 To _count - 1
                If _keys(i) = key Then Return _values(i)
            Next
            Return "Not Found"
        End Get
    End Property
    
    Default Public ReadOnly Property Item(index As Integer) As String
        Get
            If index >= 0 AndAlso index < _count Then
                Return _values(index)
            End If
            Return "Out of Bounds"
        End Get
    End Property
End Class

Module M
    Sub Main()
        Dim dict As New DictionaryMock()
        dict.Add("Apples", "Red")
        dict.Add("Bananas", "Yellow")
        
        ' Call by string key
        Console.WriteLine(dict("Bananas"))
        
        ' Call by integer index
        Console.WriteLine(dict(0))
    End Sub
End Module
