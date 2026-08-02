// vybe-test: csharp/csharp_stackalloc_span/stackalloc_span_try_copy_to_succeeds_when_dest_large_enough
// origin: languages/csharp/tests/csharp/test_csharp_stackalloc_span.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Span<int> src=stackalloc int[2]{3,4}; System.Span<int> dst=stackalloc int[3]; __Check((src.TryCopyTo(dst)).ToString(), "True");
