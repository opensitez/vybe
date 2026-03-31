Public Class SocketForm
    Private _server As Server
    Private _client As Client
    Private _serverRunning As Boolean = False

    Public Sub New()
        InitializeComponent()
        _server = New Server(8080, AddressOf LogServer)
        _client = New Client(AddressOf LogClient, AddressOf OnClientConnectionChanged)
        ' Auto-start server for verification
        _server.Start()
        _serverRunning = True
    End Sub

    Private Sub LogServer(msg As String)
        txtServerLog.AppendText("[" & DateTime.Now.ToString("HH:mm:ss") & "] " & msg & Environment.NewLine)
    End Sub

    Private Sub LogClient(msg As String)
        txtClientLog.AppendText("[" & DateTime.Now.ToString("HH:mm:ss") & "] " & msg & Environment.NewLine)
    End Sub

    Private Sub OnClientConnectionChanged(connected As Boolean)
        If connected Then
            btnConnect.Text = "Disconnect"
            btnSend.Enabled = True
        Else
            btnConnect.Text = "Connect to Server"
            btnSend.Enabled = False
        End If
    End Sub

    Private Sub btnListen_Click(sender As Object, e As EventArgs) Handles btnListen.Click
        btnListen.Text = "Are you Listening"
         If Not _serverRunning Then
           _serverRunning = True
            btnListen.Text = "Stop Listening"
            _server.Start()
        Else
            _serverRunning = False
            btnListen.Text = "Start Listening"
            _server.Stop()
        End If
    End Sub

    Private Sub btnConnect_Click(sender As Object, e As EventArgs) Handles btnConnect.Click
        If _client.Connected Then
            _client.Disconnect()
        Else
            _client.Connect("127.0.0.1", 8080)
        End If
    End Sub

    Private Sub btnSend_Click(sender As Object, e As EventArgs) Handles btnSend.Click
        Dim msg = txtMessage.Text
        _client.Send(msg)
    End Sub
End Class
