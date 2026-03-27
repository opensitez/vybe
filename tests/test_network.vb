Imports System.Net
Imports System.Net.Sockets
Imports System.IO
Imports System.Threading

Module NetworkTest
    ' Use very high ephemeral ports to avoid conflicts
    Dim basePort As Integer = 62100

    Sub Main()
        Console.WriteLine("=== Vybe Network Test ===")
        Console.WriteLine("Using base port: " & CStr(basePort))
        Console.WriteLine()

        ' Test 1: TcpListener + TcpClient basic echo
        Console.WriteLine("--- Test 1: TCP Echo (Listener + Client) ---")
        TestTcpEcho()

        ' Test 2: StreamReader/StreamWriter over TCP
        Console.WriteLine()
        Console.WriteLine("--- Test 2: StreamReader/StreamWriter over TCP ---")
        TestStreamReaderWriter()

        ' Test 3: UdpClient send/receive
        Console.WriteLine()
        Console.WriteLine("--- Test 3: UDP Send/Receive ---")
        TestUdp()

        Console.WriteLine()
        Console.WriteLine("=== All network tests complete ===")
    End Sub

    Sub TestTcpEcho()
        Dim port As Integer = basePort
        Dim listener As New TcpListener(port)
        listener.Start()
        Console.WriteLine("  Listener started on port " & CStr(port))

        ' Connect a client
        Dim client As New TcpClient("127.0.0.1", port)
        Console.WriteLine("  Client connected")

        ' Accept the connection on the server side
        Dim serverClient = listener.AcceptTcpClient()
        Console.WriteLine("  Server accepted client")

        ' Server sends a message
        Dim serverStream = serverClient.GetStream()
        Dim serverWriter As New StreamWriter(serverStream)
        serverWriter.WriteLine("Hello from server!")
        serverWriter.Flush()
        Console.WriteLine("  Server sent: Hello from server!")

        ' Client reads the message
        Dim clientStream = client.GetStream()
        Dim clientReader As New StreamReader(clientStream)
        Dim received = clientReader.ReadLine()
        Console.WriteLine("  Client received: " & received)

        ' Verify
        If received = "Hello from server!" Then
            Console.WriteLine("  PASS: TCP echo works!")
        Else
            Console.WriteLine("  FAIL: Expected 'Hello from server!', got '" & received & "'")
        End If

        ' Cleanup
        client.Close()
        serverClient.Close()
        listener.Stop()
        Console.WriteLine("  Cleanup done")
    End Sub

    Sub TestStreamReaderWriter()
        Dim port As Integer = basePort + 1
        Dim listener As New TcpListener(port)
        listener.Start()

        ' Connect client
        Dim client As New TcpClient("127.0.0.1", port)
        Dim serverClient = listener.AcceptTcpClient()

        ' Client writes via StreamWriter
        Dim clientStream = client.GetStream()
        Dim writer As New StreamWriter(clientStream)
        writer.WriteLine("Line 1")
        writer.WriteLine("Line 2")
        writer.Flush()
        Console.WriteLine("  Client sent two lines")

        ' Server reads via StreamReader
        Dim serverStream = serverClient.GetStream()
        Dim reader As New StreamReader(serverStream)
        Dim line1 = reader.ReadLine()
        Dim line2 = reader.ReadLine()
        Console.WriteLine("  Server received: '" & line1 & "' and '" & line2 & "'")

        If line1 = "Line 1" AndAlso line2 = "Line 2" Then
            Console.WriteLine("  PASS: StreamReader/StreamWriter works!")
        Else
            Console.WriteLine("  FAIL: Unexpected lines")
        End If

        ' Cleanup
        client.Close()
        serverClient.Close()
        listener.Stop()
    End Sub

    Sub TestUdp()
        Dim port1 As Integer = basePort + 2
        Dim port2 As Integer = basePort + 3
        Dim sender As New UdpClient(port1)
        Dim receiver As New UdpClient(port2)
        Console.WriteLine("  Sender on " & CStr(port1) & ", Receiver on " & CStr(port2))

        ' Send data
        Dim message = "Hello UDP!"
        sender.Send(message, Len(message), "127.0.0.1", port2)
        Console.WriteLine("  Sent: " & message)

        ' Receive data
        Dim received = receiver.Receive(Nothing)
        Console.WriteLine("  Received: " & received)

        If received = message Then
            Console.WriteLine("  PASS: UDP send/receive works!")
        Else
            Console.WriteLine("  FAIL: Expected '" & message & "', got '" & received & "'")
        End If

        ' Cleanup
        sender.Close()
        receiver.Close()
    End Sub
End Module
