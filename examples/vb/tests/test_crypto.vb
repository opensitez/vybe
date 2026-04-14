Imports System.Security.Cryptography
Imports System.Text

Module CryptoTest
    Sub Main()
        Console.WriteLine("--- Cryptography Verification ---")
        
        Dim input = "hello vybe"
        Dim bytes = Encoding.UTF8.GetBytes(input)
        
        ' 1. MD5
        Dim md5obj = MD5.Create()
        Dim hash1 = md5obj.ComputeHash(bytes)
        Dim hashStr1 = BitConverter.ToString(hash1).Replace("-", "").ToLower()
        Console.WriteLine("MD5: " & hashStr1)
        
        ' MD5 hash of "hello vybe" is b48757e77d5ada3794724f2e0c6c1d37
        If hashStr1 = "b48757e77d5ada3794724f2e0c6c1d37" Then
            Console.WriteLine("MD5 Success")
        Else
            Console.WriteLine("MD5 Failure")
        End If
        
        ' 2. SHA256
        Dim sha = SHA256.Create()
        Dim hash2 = sha.ComputeHash(bytes)
        Dim hashStr2 = BitConverter.ToString(hash2).Replace("-", "").ToLower()
        Console.WriteLine("SHA256: " & hashStr2)
        
        ' SHA256 hash of "hello vybe" is 1bb463594d86e3a7ee1e6f5712d87f65d116f5e081f9bc118d971933d16ca1da
        If hashStr2 = "1bb463594d86e3a7ee1e6f5712d87f65d116f5e081f9bc118d971933d16ca1da" Then
            Console.WriteLine("SHA256 Success")
        Else
            Console.WriteLine("SHA256 Failure")
        End If
    End Sub
End Module
