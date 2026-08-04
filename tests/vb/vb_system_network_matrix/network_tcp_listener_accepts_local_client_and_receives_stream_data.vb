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
'
' Output is COLLECTED, not paired. The emitter rewrites every
' `Console.WriteLine(x)` into `__P(CStr(x))` and compares the whole output once
' at the end of `Sub Main`. Pairing the i-th print with the i-th expected line
' cannot assert anything about a loop, and loops alone were 402 of VB's 6,671
' cases.
'
' Rendering happens at the CALL SITE via `CStr`, where the expression still has
' its static type — the same reason the C# harness renders with `.ToString()`
' rather than inside the helper.

Module VybeCheck
    Public __buf As String = ""

    Sub __P(s As String)
        __buf = __buf & s & vbLf
    End Sub

    Sub __Pr(s As String)
        __buf = __buf & s
    End Sub

    ' The final WriteLine contributes a trailing newline that the expected line
    ' vector never carried, so BOTH forms are accepted.
    Sub __Check(want As String)
        If __buf <> want AndAlso __buf <> want & vbLf Then
            Console.WriteLine("FAIL: want [" & want & "] got [" & __buf & "]")
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
                __Check("ok
True
True
True")
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

        __P(CStr(Encoding.UTF8.GetString(response, 0, count)))
        __P(CStr(count = 2))
        __P(CStr(handshakeDone))
        __P(CStr(Not serverThread.IsAlive))
    End Sub
End Module
