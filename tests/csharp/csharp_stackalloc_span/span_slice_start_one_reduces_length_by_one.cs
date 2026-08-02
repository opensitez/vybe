// vybe-test: csharp/csharp_stackalloc_span/span_slice_start_one_reduces_length_by_one
// origin: languages/csharp/tests/csharp/test_csharp_stackalloc_span.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Span<int> span=stackalloc int[4]{1,2,3,4}; var tail=span.Slice(1); __Check((tail.Length).ToString(), "3");
