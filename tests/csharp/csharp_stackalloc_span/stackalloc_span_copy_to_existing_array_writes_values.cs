// vybe-test: csharp/csharp_stackalloc_span/stackalloc_span_copy_to_existing_array_writes_values
// origin: languages/csharp/tests/csharp/test_csharp_stackalloc_span.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Span<int> src=stackalloc int[2]{8,9}; int[] dst=new int[2]; src.CopyTo(dst); __Check((dst[1]).ToString(), "9");
