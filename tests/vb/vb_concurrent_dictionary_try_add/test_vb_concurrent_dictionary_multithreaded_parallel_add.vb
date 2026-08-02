' vybe-test: vb/vb_concurrent_dictionary_try_add/test_vb_concurrent_dictionary_multithreaded_parallel_add
' origin: languages/vb/tests/vb/test_vb_concurrent_dictionary_try_add.rs

Imports System.Collections.Concurrent
Imports System.Threading.Tasks

Module Program
    Sub Main()
        Dim dict As New ConcurrentDictionary(Of Integer, Integer)()
        Parallel.For(0, 100, Sub(i) dict.TryAdd(i, i * 2))
        Console.WriteLine("Total Count: " & dict.Count)
    End Sub
End Module
