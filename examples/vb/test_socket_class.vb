Public Class Server
    Private _port As Integer
    Private _running As Boolean = False
    Private _logCallback As Action(Of String)

    Public Sub New(port As Integer, logCallback As Action(Of String))
        _port = port
        _logCallback = logCallback
        Console.WriteLine("Server created on port " & _port)
    End Sub

    Public Sub Start()
        If _running Then Return
        _running = True
        Console.WriteLine("Server started on port " & _port)
    End Sub

    Public Sub [Stop]()
        _running = False
        Console.WriteLine("Server stopped")
    End Sub
End Class

Public Class TestForm
    Private _server As Server
    Private _serverRunning As Boolean = False

    Public Sub New()
        Console.WriteLine("TestForm constructor")
        _server = New Server(8080, AddressOf LogServer)
        _server.Start()
        _serverRunning = True
        Console.WriteLine("Constructor done, _serverRunning = " & _serverRunning)
    End Sub

    Private Sub LogServer(msg As String)
        Console.WriteLine("[Server] " & msg)
    End Sub

    Public Sub btnListen_Click()
        Console.WriteLine("Click handler fired, _serverRunning = " & _serverRunning)
        If Not _serverRunning Then
            _serverRunning = True
            Console.WriteLine("Started listening")
            _server.Start()
        Else
            _serverRunning = False
            Console.WriteLine("Stopped listening")
            _server.Stop()
        End If
        Console.WriteLine("Click handler done, _serverRunning = " & _serverRunning)
    End Sub
End Class

Module Module1
    Sub Main()
        Dim form As New TestForm()
        Console.WriteLine("--- Click 1 ---")
        form.btnListen_Click()
        Console.WriteLine("--- Click 2 ---")
        form.btnListen_Click()
    End Sub
End Module
