// vybe-test: csharp/csharp_stackalloc_span/stackalloc_double_buffer_reads_fractional_value
// origin: languages/csharp/tests/csharp/test_csharp_stackalloc_span.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

System.Span<double> buf=stackalloc double[2]{1.5,2.5}; __Check((buf[1]).ToString(), "2.5");
