Imports System
Imports System.Net

Module Program
    Sub Main()
        Console.WriteLine("=== System.Net Demo ===")
        Console.WriteLine()

        ' --- 1. WebClient.DownloadString ---
        Console.WriteLine("--- WebClient.DownloadString ---")
        Dim client As New WebClient()
        Dim result As String = client.DownloadString("https://httpbin.org/ip")
        Console.WriteLine(result)

        ' --- 2. WebClient.DownloadFile ---
        Console.WriteLine("--- WebClient.DownloadFile ---")
        client.DownloadFile("https://httpbin.org/robots.txt", "/tmp/robots.txt")
        Console.WriteLine("Downloaded to /tmp/robots.txt")

        ' --- 3. WebClient with Headers ---
        Console.WriteLine("--- WebClient with Headers ---")
        Dim client2 As New WebClient()
        client2.Headers.Add("X-Custom-Header", "vybeTest")
        client2.Headers.Add("Accept", "application/json")
        Dim headerResult As String = client2.DownloadString("https://httpbin.org/headers")
        Console.WriteLine(headerResult)

        ' --- 4. WebClient.UploadString (POST) ---
        Console.WriteLine("--- WebClient.UploadString ---")
        Dim postData As String = "{""name"":""vybe""}"
        Dim postResult As String = client.UploadString("https://httpbin.org/post", postData)
        Console.WriteLine("POST response length: " & postResult.Length)

        ' --- 5. HttpClient.GetStringAsync ---
        Console.WriteLine("--- HttpClient.GetStringAsync ---")
        Dim http As New System.Net.Http.HttpClient()
        Dim body As String = http.GetStringAsync("https://httpbin.org/user-agent")
        Console.WriteLine(body)

        ' --- 6. HttpClient.GetAsync (with response object) ---
        Console.WriteLine("--- HttpClient.GetAsync ---")
        Dim resp = http.GetAsync("https://httpbin.org/status/200")
        Console.WriteLine("Status: " & resp.StatusCode)
        Console.WriteLine("Success: " & resp.IsSuccessStatusCode)

        ' --- 7. Dns.GetHostEntry ---
        Console.WriteLine("--- Dns.GetHostEntry ---")
        Dim entry = Dns.GetHostEntry("example.com")
        Console.WriteLine("Hostname: " & entry.HostName)
        Dim addrs() As String = entry.AddressList
        Dim i As Integer
        For i = 0 To addrs.Length - 1
            Console.WriteLine("  Address: " & addrs(i))
        Next

        ' --- 8. Dns.GetHostName ---
        Console.WriteLine("--- Dns.GetHostName ---")
        Console.WriteLine("Local hostname: " & Dns.GetHostName())

        ' --- 9. IPAddress.Parse ---
        Console.WriteLine("--- IPAddress.Parse ---")
        Dim ip = IPAddress.Parse("192.168.1.1")
        Console.WriteLine("Address: " & ip.Address)
        Console.WriteLine("Family: " & ip.AddressFamily)

        Dim ip6 = IPAddress.Parse("::1")
        Console.WriteLine("IPv6: " & ip6.Address & " (" & ip6.AddressFamily & ")")

        ' --- 10. IPAddress.TryParse ---
        Console.WriteLine("--- IPAddress.TryParse ---")
        Dim valid As Boolean = IPAddress.TryParse("10.0.0.1")
        Console.WriteLine("10.0.0.1 valid: " & valid)
        Dim invalid As Boolean = IPAddress.TryParse("not.an.ip")
        Console.WriteLine("not.an.ip valid: " & invalid)

        ' --- 11. IPAddress constants ---
        Console.WriteLine("--- IPAddress Constants ---")
        Dim lo = IPAddress.Loopback
        Console.WriteLine("Loopback: " & lo.Address)
        Dim any = IPAddress.Any
        Console.WriteLine("Any: " & any.Address)

        Console.WriteLine()
        Console.WriteLine("=== All tests complete ===")
    End Sub
End Module
