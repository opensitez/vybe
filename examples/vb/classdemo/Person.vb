Public Class Person
    ' Private field (data storage)
    Private _name As String

    ' Public Property
    Public Property Name() As String
        Get
            Return _name
        End Get
        Set(ByVal value As String)
            _name = value
        End Set
    End Property

    ' A simple Method (Action)
    Public Sub Greet()
        MessageBox.Show("Hello, my name is " & _name, "Greeting")
    End Sub
End Class
