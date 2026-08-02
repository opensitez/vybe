// vybe-test: csharp/csharp_stackalloc_span/span_slice_start_and_length_reads_subrange
// origin: languages/csharp/tests/csharp/test_csharp_stackalloc_span.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Span<int> span=stackalloc int[5]{10,20,30,40,50}; var mid=span.Slice(1,2); __Check((mid[0]).ToString(), "20"); __Check((mid[1]).ToString(), "30");
