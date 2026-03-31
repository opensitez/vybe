Imports System.Net
Imports System.Net.Sockets
Imports System.IO
Imports System.Threading.Tasks

Public Class Server
    Private _listener As TcpListener
    Private _port As Integer
    Private _running As Boolean = False
    Private _logCallback As Action(Of String)

    Public Sub New(port As Integer, logCallback As Action(Of String))
        _port = port
        _logCallback = logCallback
    End Sub

    Public Sub Start()
        If _running Then Return
        _running = True
        
        Task.Run(Sub()
            Try
                _listener = New TcpListener(_port)
                _listener.Start()
                _logCallback("Server started on port " & _port)

                While _running
                    If _listener.Pending() Then
                        Dim remoteClient = _listener.AcceptTcpClient()
                        _logCallback("Client connected: " & remoteClient.ToString())
                        
                        Task.Run(Sub()
                            HandleClient(remoteClient)
                        End Sub)
                    End If
                    Threading.Thread.Sleep(100)
                End While

                _listener.Stop()
                _logCallback("Server stopped.")
            Catch ex As Exception
                _logCallback("Server Error: " & ex.Message)
                _running = False
            End Try
        End Sub)
    End Sub

    Public Sub [Stop]()
        _running = False
    End Sub

    Private Sub HandleClient(client As TcpClient)
        Try
            Dim stream = client.GetStream()
            Dim reader As New StreamReader(stream)
            Dim writer As New StreamWriter(stream)

            While _running AndAlso client.Connected
                Dim receivedText = reader.ReadLine()
                If receivedText IsNot Nothing Then
                    _logCallback("Received: " & receivedText)
                    writer.WriteLine("Echo: " & receivedText)
                    writer.Flush()
                Else
                    Exit While
                End If
            End While
            client.Close()
            _logCallback("Client disconnected.")
        Catch ex As Exception
            _logCallback("HandleClient Error: " & ex.Message)
        End Try
    End Sub
End Class
