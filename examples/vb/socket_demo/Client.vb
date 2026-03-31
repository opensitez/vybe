Imports System.Net
Imports System.Net.Sockets
Imports System.IO
Imports System.Threading.Tasks

Public Class Client
    Private _client As TcpClient
    Private _writer As StreamWriter
    Private _reader As StreamReader
    Private _logCallback As Action(Of String)
    Private _connectedCallback As Action(Of Boolean)

    Public Sub New(logCallback As Action(Of String), connectedCallback As Action(Of Boolean))
        _logCallback = logCallback
        _connectedCallback = connectedCallback
    End Sub

    Public ReadOnly Property Connected As Boolean
        Get
            Return _client IsNot Nothing AndAlso _client.Connected
        End Get
    End Property

    Public Sub Connect(host As String, port As Integer)
        Try
            _client = New TcpClient(host, port)
            Dim stream = _client.GetStream()
            _writer = New StreamWriter(stream)
            _reader = New StreamReader(stream)

            _logCallback("Connected to " & host & ":" & port)
            _connectedCallback(True)

            ' Start background read loop
            Task.Run(Sub()
                Try
                    While _client IsNot Nothing AndAlso _client.Connected
                        Dim receivedText = _reader.ReadLine()
                        If receivedText IsNot Nothing Then
                            _logCallback("Server Says: " & receivedText)
                        Else
                            Exit While
                        End If
                    End While
                    _logCallback("Connection closed by server.")
                Catch ex As Exception
                    _logCallback("Read Loop Error: " & ex.Message)
                End Try
                _connectedCallback(False)
            End Sub)

        Catch ex As Exception
            _logCallback("Connection Failed: " & ex.Message)
            _connectedCallback(False)
        End Try
    End Sub

    Public Sub Send(msg As String)
        Try
            If Connected Then
                _writer.WriteLine(msg)
                _writer.Flush()
                _logCallback("Sent: " & msg)
            End If
        Catch ex As Exception
            _logCallback("Send Error: " & ex.Message)
        End Try
    End Sub

    Public Sub Disconnect()
        If _client IsNot Nothing Then
            _client.Close()
            _client = Nothing
            _logCallback("Disconnected.")
        End If
    End Sub
End Class
