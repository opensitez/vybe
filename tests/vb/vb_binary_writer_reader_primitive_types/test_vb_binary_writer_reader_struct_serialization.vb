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
                __P(CStr(restored.ID & ":" & restored.Name & "=" & restored.Score))
            End Using
        End Using
        __Check("1:Hero=95.5")
    End Sub
End Module
