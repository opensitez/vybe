' vybe-test: vb/vb_struct_layout_sequential/test_vb_struct_layout_auto_layout_kind
' origin: languages/vb/tests/vb/test_vb_struct_layout_sequential.rs

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

Imports System.Runtime.InteropServices

<StructLayout(LayoutKind.Auto)>
Class AutoLayoutClass
    Public A As Byte
    Public B As Integer
End Class

Module Program
    Sub Main()
        ' Auto layout classes cannot be measured via Marshal.SizeOf directly without throwing!
        Try
            Marshal.SizeOf(GetType(AutoLayoutClass))
        Catch ex As ArgumentException
            __Check(CStr("ArgumentException Caught on Auto Layout"), "ArgumentException Caught on Auto Layout")
        End Try
    End Sub
End Module
