' vybe-test: vb/vb_extension_methods_advanced/extension_methods_generics
' origin: languages/vb/tests/vb/test_vb_extension_methods_advanced.rs

Imports System.Runtime.CompilerServices
Imports System.Collections.Generic

Module ExtensionMethods
    <Extension()>
    Public Sub PrintItems(Of T)(collection As IEnumerable(Of T))
        For Each item In collection
            Console.WriteLine(item)
        Next
    End Sub
End Module

Module M
    Sub Main()
        Dim list As New List(Of Integer) From { 1, 2, 3 }
        list.PrintItems()
    End Sub
End Module
