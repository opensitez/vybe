' vybe-test: vb/vb_for_each_custom_enumerator_struct/test_vb_for_each_exception_during_enumeration
' origin: languages/vb/tests/vb/test_vb_for_each_custom_enumerator_struct.rs

Imports System
Imports System.Collections.Generic

Module Program
    Private Iterator Function FaultyEnum() As IEnumerable(Of Integer)
        Yield 1
        Throw New InvalidOperationException("Faulty Enum Error")
    End Function

    Sub Main()
        Try
            For Each item In FaultyEnum()
                Console.WriteLine(item)
            Next
        Catch ex As InvalidOperationException
            Console.WriteLine(ex.Message)
        End Try
    End Sub
End Module
