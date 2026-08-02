// vybe-test: csharp/csharp_stackalloc_span/stackalloc_span_try_copy_to_fails_when_dest_too_small
// origin: languages/csharp/tests/csharp/test_csharp_stackalloc_span.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Span<int> src=stackalloc int[3]{1,2,3}; System.Span<int> dst=stackalloc int[2]; __Check((src.TryCopyTo(dst)).ToString(), "False");
