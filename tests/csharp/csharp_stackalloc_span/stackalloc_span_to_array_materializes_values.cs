// vybe-test: csharp/csharp_stackalloc_span/stackalloc_span_to_array_materializes_values
// origin: languages/csharp/tests/csharp/test_csharp_stackalloc_span.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Span<int> span=stackalloc int[2]{12,34}; int[] arr=span.ToArray(); __Check((arr[1]).ToString(), "34");
