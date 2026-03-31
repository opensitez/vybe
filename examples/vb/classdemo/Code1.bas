Attribute VB_Name = "Code1"
' Code file: Code1

Public Class Person
    ' Private field (data storage)
    Private _name As String

    ' Public Property (The "Visual Basic" way to access data)
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
        MsgBox("Hello, my name is " & _name)
    End Sub
End Class
