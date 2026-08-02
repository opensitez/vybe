' vybe-test: vb/vb_interface_default_methods_adv/test_vb_interface_implements_multiple_interfaces
' origin: languages/vb/tests/vb/test_vb_interface_default_methods_adv.rs

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

Interface IReader
    Function ReadData() As String
End Interface

Interface IWriter
    Sub WriteData(data As String)
End Interface

Class StorageStream
    Implements IReader, IWriter
    Private buffer As String = ""
    Public Function ReadData() As String Implements IReader.ReadData
        Return buffer
    End Function
    Public Sub WriteData(data As String) Implements IWriter.WriteData
        buffer = data
    End Sub
End Class

Module Program
    Sub Main()
        Dim s As New StorageStream()
        Dim w As IWriter = s
        w.WriteData("Hello World")
        Dim r As IReader = s
        __Check(CStr(r.ReadData()), "Hello World")
    End Sub
End Module
