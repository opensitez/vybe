// vybe-test: csharp/csharp_stackalloc_span/stackalloc_span_slice_empty_at_end_has_zero_length
// origin: languages/csharp/tests/csharp/test_csharp_stackalloc_span.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Span<int> span=stackalloc int[2]{1,2}; var empty=span.Slice(2); __Check((empty.Length).ToString(), "0");
