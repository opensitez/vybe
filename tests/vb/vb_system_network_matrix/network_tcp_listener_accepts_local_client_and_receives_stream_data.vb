' vybe-test: vb/vb_system_network_matrix/network_tcp_listener_accepts_local_client_and_receives_stream_data
' origin: languages/vb/tests/vb/test_vb_system_network_matrix.rs

' Vybe test harness — Visual Basic.
'
' Real VB source alongside harness/go/check.go and harness/js/check.js, the way
' test262's assert.js is JavaScript.
'
' A test's verdict is its EXIT CODE. __Check prints its diagnostic BEFORE
' throwing: an uncaught exception surfaces as `RuntimeError: [object]`, which
' says nothing at all.

Module VybeCheck
    Sub __Check(got As String, want As String)
        If got <> want Then
            Console.WriteLine("FAIL: want [" & want & "] got [" & got & "]")
            Throw New Exception("assertion failed")
        End If
    End Sub
End Module

Imports System.Net
Imports System.Net.Sockets
Imports System.Text
Imports System.Threading

Module M
    Sub Main()
        Dim listener As New TcpListener(0)
        listener.Start()

        Dim port As Integer = CType(listener.LocalEndpoint, IPEndPoint).Port
        Dim reply() As Byte = Encoding.UTF8.GetBytes("ok")
        Dim handshakeDone As Boolean = False
        Dim serverThread As New Thread(
            Sub()
                Dim accepted As TcpClient = listener.AcceptTcpClient()
                Dim outStream As NetworkStream = accepted.GetStream()
                outStream.Write(reply, 0, reply.Length)
                outStream.Flush()
                accepted.Close()
                handshakeDone = True
            End Sub
        )

        serverThread.Start()

        Dim client As New TcpClient()
        client.Connect("127.0.0.1", port)
        Dim clientStream As NetworkStream = client.GetStream()
        Dim response(1) As Byte
        Dim count As Integer = clientStream.Read(response, 0, response.Length)
        client.Close()
        listener.Stop()
        serverThread.Join(1000)

        __Check(CStr(Encoding.UTF8.GetString(response, 0, count)), "ok")
        __Check(CStr(count = 2), "True")
        __Check(CStr(handshakeDone), "True")
        __Check(CStr(Not serverThread.IsAlive), "True")
    End Sub
End Module
