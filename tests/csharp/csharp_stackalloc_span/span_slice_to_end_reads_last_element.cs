// vybe-test: csharp/csharp_stackalloc_span/span_slice_to_end_reads_last_element
// origin: languages/csharp/tests/csharp/test_csharp_stackalloc_span.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Span<int> span=stackalloc int[3]{5,6,7}; var last=span.Slice(2); __Check((last[0]).ToString(), "7");
