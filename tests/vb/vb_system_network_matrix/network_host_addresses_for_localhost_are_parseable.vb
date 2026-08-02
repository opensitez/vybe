' vybe-test: vb/vb_system_network_matrix/network_host_addresses_for_localhost_are_parseable
' origin: languages/vb/tests/vb/test_vb_system_network_matrix.rs

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
