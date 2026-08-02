' vybe-test: vb/vb_binary_writer_reader_primitive_types/test_vb_binary_writer_reader_struct_serialization
' origin: languages/vb/tests/vb/test_vb_binary_writer_reader_primitive_types.rs

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

Imports System.IO

Structure PlayerState
    Public ID As Integer
    Public Name As String
    Public Score As Double
End Structure

Module Program
    Sub Main()
        Dim player As New PlayerState With {.ID = 1, .Name = "Hero", .Score = 95.5}
        Using ms As New MemoryStream()
            Using bw As New BinaryWriter(ms, System.Text.Encoding.UTF8, True)
                bw.Write(player.ID)
                bw.Write(player.Name)
                bw.Write(player.Score)
            End Using

            ms.Position = 0
            Using br As New BinaryReader(ms)
                Dim restored As New PlayerState With {
                    .ID = br.ReadInt32(),
                    .Name = br.ReadString(),
                    .Score = br.ReadDouble()
                }
                __Check(CStr(restored.ID & ":" & restored.Name & "=" & restored.Score), "1:Hero=95.5")
            End Using
        End Using
    End Sub
End Module
