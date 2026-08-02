// vybe-test: csharp/csharp_stackalloc_span/stackalloc_int_zero_length_span_has_zero_length
// origin: languages/csharp/tests/csharp/test_csharp_stackalloc_span.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Span<int> buf=stackalloc int[0]; __Check((buf.Length).ToString(), "0");
