' vybe-test: vb/vb_generic_interface_inheritance/test_vb_generic_interface_struct_implementer
' origin: languages/vb/tests/vb/test_vb_generic_interface_inheritance.rs

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

Interface ISwap(Of T)
    Function SwapWith(other As T) As T
End Interface

Structure Pair(Of T)
    Implements ISwap(Of Pair(Of T))
    Public First As T
    Public Second As T
    Public Sub New(f As T, s As T)
        First = f : Second = s
    End Sub
    Public Function SwapWith(other As Pair(Of T)) As Pair(Of T) Implements ISwap(Of Pair(Of T)).SwapWith
        Return New Pair(Of T)(Second, First)
    End Function
End Structure

Module Program
    Sub Main()
        Dim p As New Pair(Of Integer)(10, 20)
        Dim swapped = p.SwapWith(p)
        __Check(CStr(swapped.First & "," & swapped.Second), "20,10")
    End Sub
End Module
