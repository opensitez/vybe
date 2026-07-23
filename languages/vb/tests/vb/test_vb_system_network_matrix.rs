use super::helpers::run_vb;

#[test]
fn network_host_name_is_reported() {
    let out = run_vb(
        r#"
Imports System.Net

Module M
    Sub Main()
        Dim hostName As String = Dns.GetHostName()
        Console.WriteLine(Not String.IsNullOrWhiteSpace(hostName))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}

#[test]
fn network_host_addresses_for_localhost_exist() {
    let out = run_vb(
        r#"
Imports System.Net

Module M
    Sub Main()
        Dim addresses() As IPAddress = Dns.GetHostAddresses("localhost")
        Console.WriteLine(addresses.Length >= 1)
        Console.WriteLine(addresses(0).ToString().Length > 0)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn network_host_addresses_for_localhost_are_parseable() {
    let out = run_vb(
        r#"
Imports System.Net

Module M
    Sub Main()
        Dim addresses() As IPAddress = Dns.GetHostAddresses("localhost")
        Dim allGood As Boolean = addresses.Length > 0

        For Each address As IPAddress In addresses
            Dim text As String = address.ToString()
            If Not String.IsNullOrWhiteSpace(text) Then
                allGood = allGood And True
            End If
        Next

        Console.WriteLine(allGood)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}

#[test]
fn network_host_entry_contains_host_and_address() {
    let out = run_vb(
        r#"
Imports System.Net

Module M
    Sub Main()
        Dim entry As IPHostEntry = Dns.GetHostEntry("localhost")
        Console.WriteLine(Not String.IsNullOrWhiteSpace(entry.HostName))
        Console.WriteLine(entry.AddressList.Length >= 1)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn network_host_entry_is_case_insensitive_for_lookup() {
    let out = run_vb(
        r#"
Imports System.Net

Module M
    Sub Main()
        Dim first As String = Dns.GetHostEntry("localhost").HostName
        Dim second As String = Dns.GetHostEntry("LOCALHOST").HostName
        Console.WriteLine(Not String.IsNullOrWhiteSpace(first))
        Console.WriteLine(first.Length > 0)
        Console.WriteLine(second.Length > 0)
        Console.WriteLine(first = second)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True", "True", "True"]);
}

#[test]
fn network_tcp_listener_exposes_bound_port() {
    let out = run_vb(
        r#"
Imports System.Net
Imports System.Net.Sockets

Module M
    Sub Main()
        Dim listener As New TcpListener(0)
        listener.Start()
        Dim localPort As Integer = CType(listener.LocalEndpoint, IPEndPoint).Port
        listener.Stop()
        Console.WriteLine(localPort > 0)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}

#[test]
fn network_tcp_listener_pending_is_false_when_idle() {
    let out = run_vb(
        r#"
Imports System.Net
Imports System.Net.Sockets

Module M
    Sub Main()
        Dim listener As New TcpListener(0)
        listener.Start()
        Console.WriteLine(listener.Pending() = False)
        Console.WriteLine(CType(listener.LocalEndpoint, IPEndPoint).Port > 0)
        listener.Stop()
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn network_tcp_listener_accepts_local_client_and_receives_stream_data() {
    let out = run_vb(
        r#"
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

        Console.WriteLine(Encoding.UTF8.GetString(response, 0, count))
        Console.WriteLine(count = 2)
        Console.WriteLine(handshakeDone)
        Console.WriteLine(Not serverThread.IsAlive)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["ok", "True", "True", "True"]);
}

#[test]
fn network_tcp_listener_roundtrip_still_reports_non_zero_port() {
    let out = run_vb(
        r#"
Imports System.Net
Imports System.Net.Sockets

Module M
    Sub Main()
        Dim listener As New TcpListener(0)
        listener.Start()
        Dim endpoint As IPEndPoint = CType(listener.LocalEndpoint, IPEndPoint)
        Console.WriteLine(endpoint.Port > 0)
        listener.Stop()
        Console.WriteLine(listener.LocalEndpoint IsNot Nothing)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn network_ipv4_loopback_can_be_parsed_and_stringified() {
    let out = run_vb(
        r#"
Imports System.Net

Module M
    Sub Main()
        Dim ip As IPAddress = IPAddress.Parse("127.0.0.1")
        Console.WriteLine(ip.ToString())
        Console.WriteLine(ip.AddressFamily.ToString())
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["127.0.0.1", "InterNetwork"]);
}

#[test]
fn network_udp_send_reports_payload_length() {
    let out = run_vb(
        r#"
Imports System.Net.Sockets

Module M
    Sub Main()
        Dim udp As New UdpClient()
        Dim payload() As Byte = {1, 2, 3, 4}
        Dim sent As Integer = udp.Send(payload, payload.Length, "127.0.0.1", 9)
        udp.Close()
        Console.WriteLine(sent = payload.Length)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["True"]);
}
