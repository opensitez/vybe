' vybe-test: vb/vb_unsafe_pointer_arithmetic/test_vb_pointer_structure_array_indexing
' origin: languages/vb/tests/vb/test_vb_unsafe_pointer_arithmetic.rs

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

Imports System
Imports System.Runtime.InteropServices

<StructLayout(LayoutKind.Sequential)>
Structure Element
    Public ID As Integer
    Public Value As Double
End Structure

Module Program
    Sub Main()
        Dim elements(1) As Element
        elements(0) = New Element With {.ID = 1, .Value = 1.1}
        elements(1) = New Element With {.ID = 2, .Value = 2.2}

        Dim handle = GCHandle.Alloc(elements, GCHandleType.Pinned)
        Dim baseAddr = handle.AddrOfPinnedObject()

        Dim elemSize = Marshal.SizeOf(GetType(Element))
        Dim secondElemAddr = IntPtr.Add(baseAddr, elemSize)
        Dim restored2 As Element = CType(Marshal.PtrToStructure(secondElemAddr, GetType(Element)), Element)
        handle.Free()

        __Check(CStr(restored2.ID & ":" & restored2.Value), "2:2.2")
    End Sub
End Module
